'use client';

import { useEffect, useRef, useState, useCallback } from 'react';
import { Mic, Square, FileSpreadsheet, Loader2, AlertCircle, CheckCircle2 } from 'lucide-react';
import { supabase } from '@/lib/supabase';
import type { TranscriptEntry, SessionInfo } from '@/lib/types';

// 録音フェーズの型定義
type Phase = 'validating' | 'ready' | 'recording' | 'finalizing' | 'error';

interface Props {
  sessionId: string;
  token: string;
}

export default function RecordingConsole({ sessionId, token }: Props) {
  // ---- 状態管理 ----
  const [phase, setPhase] = useState<Phase>('validating');
  const [session, setSession] = useState<SessionInfo | null>(null);
  const [transcripts, setTranscripts] = useState<TranscriptEntry[]>([]);
  // 未確定（interim）のトランスクリプトエントリ
  const [pendingTranscript, setPendingTranscript] = useState<TranscriptEntry | null>(null);
  const [error, setError] = useState<string | null>(null);
  // 録音経過時間（秒）
  const [elapsedSeconds, setElapsedSeconds] = useState(0);

  // ---- Ref（再レンダリングをまたいで保持する値）----
  const wsRef = useRef<WebSocket | null>(null);
  const audioContextRef = useRef<AudioContext | null>(null);
  const workletNodeRef = useRef<AudioWorkletNode | null>(null);
  const mediaStreamRef = useRef<MediaStream | null>(null);
  const timerRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const transcriptEndRef = useRef<HTMLDivElement | null>(null);

  // ---- トランスクリプト末尾への自動スクロール ----
  useEffect(() => {
    transcriptEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [transcripts, pendingTranscript]);

  // ---- マウント時: トークン検証 ----
  useEffect(() => {
    const validate = async () => {
      // sessionStorage に認証済みキーがあればトークン検証をスキップ（リロード対応）
      const storageKey = `authorized:${sessionId}`;
      const alreadyAuthorized = sessionStorage.getItem(storageKey);

      if (!alreadyAuthorized) {
        // トークンが未指定の場合はエラー
        if (!token) {
          setError('アクセストークンが指定されていません。SonaCoreから発行されたリンクを使用してください。');
          setPhase('error');
          return;
        }

        // launch_tokens テーブルでトークンを検証
        const now = new Date().toISOString();
        const { data: tokenRow, error: tokenError } = await supabase
          .from('launch_tokens')
          .select('id, session_id')
          .eq('session_id', sessionId)
          .eq('token', token)
          .is('used_at', null)
          .gt('expires_at', now)
          .single();

        if (tokenError || !tokenRow) {
          setError('トークンが無効または期限切れです。SonaCoreから新しいリンクを発行してください。');
          setPhase('error');
          return;
        }

        // トークンを消費（used_at を現在時刻で更新）
        const { error: updateError } = await supabase
          .from('launch_tokens')
          .update({ used_at: now })
          .eq('id', tokenRow.id);

        if (updateError) {
          setError('トークンの消費に失敗しました。再度お試しください。');
          setPhase('error');
          return;
        }

        // sessionStorage に認証済みフラグを保存
        sessionStorage.setItem(storageKey, '1');
      }

      // セッション情報を取得
      const { data: sessionData, error: sessionError } = await supabase
        .from('sessions')
        .select('id, name, status, description')
        .eq('id', sessionId)
        .single();

      if (sessionError || !sessionData) {
        setError('セッション情報の取得に失敗しました。');
        setPhase('error');
        return;
      }

      setSession(sessionData as SessionInfo);
      setPhase('ready');
    };

    validate();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ---- 録音開始 ----
  const handleStartRecording = useCallback(async () => {
    try {
      // マイクへのアクセス許可を要求
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      mediaStreamRef.current = stream;

      // 16kHz の AudioContext を作成
      const audioContext = new AudioContext({ sampleRate: 16000 });
      audioContextRef.current = audioContext;

      // AudioWorklet モジュールを読み込む
      await audioContext.audioWorklet.addModule('/audio-processor.js');

      // マイク入力ソースを作成
      const source = audioContext.createMediaStreamSource(stream);

      // PCM 変換ワークレットノードを作成
      const workletNode = new AudioWorkletNode(audioContext, 'pcm-processor');
      workletNodeRef.current = workletNode;

      // WebSocket 接続を確立
      const wsUrl = process.env.NEXT_PUBLIC_VOICE_VERIFIER_WS_URL ?? 'ws://localhost:8000';
      const ws = new WebSocket(`${wsUrl}/ws/transcribe`);
      wsRef.current = ws;
      ws.binaryType = 'arraybuffer';

      ws.onopen = () => {
        // 接続確立後にセッション設定を送信
        ws.send(JSON.stringify({ type: 'config', session_id: sessionId }));
      };

      ws.onmessage = (event) => {
        try {
          const data = JSON.parse(event.data as string) as {
            type: string;
            text: string;
            speaker: string;
            is_final: boolean;
          };

          if (data.type !== 'transcript') return;

          const entry: TranscriptEntry = {
            id: `${Date.now()}-${Math.random()}`,
            text: data.text,
            speaker: data.speaker,
            isFinal: data.is_final,
            timestamp: new Date().toLocaleTimeString('ja-JP'),
          };

          if (data.is_final) {
            // 確定済みエントリをリストに追加し、未確定をクリア
            setTranscripts((prev) => [...prev, entry]);
            setPendingTranscript(null);
          } else {
            // 未確定エントリを上書き更新
            setPendingTranscript(entry);
          }
        } catch {
          // JSON パース失敗は無視
        }
      };

      ws.onerror = () => {
        setError('WebSocket 接続でエラーが発生しました。');
      };

      // ワークレットから PCM バイナリを受け取り WebSocket へ送信
      workletNode.port.onmessage = (event: MessageEvent<ArrayBuffer>) => {
        if (ws.readyState === WebSocket.OPEN) {
          ws.send(event.data);
        }
      };

      // マイク → ワークレットノード → destination に接続
      source.connect(workletNode);
      workletNode.connect(audioContext.destination);

      // 経過時間カウンターを開始
      setElapsedSeconds(0);
      timerRef.current = setInterval(() => {
        setElapsedSeconds((s) => s + 1);
      }, 1000);

      setPhase('recording');
    } catch (err) {
      const message = err instanceof Error ? err.message : '不明なエラー';
      setError(`録音の開始に失敗しました: ${message}`);
      setPhase('error');
    }
  }, [sessionId]);

  // ---- 録音停止 ----
  const handleStopRecording = useCallback(async () => {
    setPhase('finalizing');

    // タイマーを停止
    if (timerRef.current) {
      clearInterval(timerRef.current);
      timerRef.current = null;
    }

    // AudioContext と AudioWorkletNode をクリーンアップ
    workletNodeRef.current?.disconnect();
    workletNodeRef.current = null;

    if (audioContextRef.current) {
      await audioContextRef.current.close();
      audioContextRef.current = null;
    }

    // マイクストリームを停止
    mediaStreamRef.current?.getTracks().forEach((track) => track.stop());
    mediaStreamRef.current = null;

    // WebSocket を閉じる
    if (wsRef.current) {
      wsRef.current.close();
      wsRef.current = null;
    }

    // セッションのステータスを 'completed' に更新
    await supabase
      .from('sessions')
      .update({ status: 'completed' })
      .eq('id', sessionId);

    // トランスクリプトデータを session_agendas テーブルに保存
    if (transcripts.length > 0) {
      await supabase.from('session_agendas').insert({
        session_id: sessionId,
        agenda_data: transcripts,
        created_at: new Date().toISOString(),
      });
    }

    // セッション情報のステータスをローカルでも更新
    setSession((prev) => (prev ? { ...prev, status: 'completed' } : prev));
    setPhase('ready');
  }, [sessionId, transcripts]);

  // ---- Excel 出力 ----
  const handleExport = useCallback(async () => {
    try {
      const sonaWebUrl = process.env.NEXT_PUBLIC_SONA_WEB_URL ?? 'http://localhost:3000';
      const response = await fetch(`${sonaWebUrl}/api/export/excel`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sessionId }),
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      // レスポンスをバイナリとして受け取り、ダウンロードする
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `session_${sessionId}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      const message = err instanceof Error ? err.message : '不明なエラー';
      alert(`Excel 出力に失敗しました: ${message}`);
    }
  }, [sessionId]);

  // ---- 経過時間のフォーマット（MM:SS）----
  const formatElapsed = (seconds: number) => {
    const m = Math.floor(seconds / 60).toString().padStart(2, '0');
    const s = (seconds % 60).toString().padStart(2, '0');
    return `${m}:${s}`;
  };

  // ---- ステータスバッジの色 ----
  const statusBadgeClass = (status: string) => {
    switch (status) {
      case 'completed':
        return 'bg-blue-900 text-blue-300';
      case 'recording':
        return 'bg-red-900 text-red-300';
      case 'active':
        return 'bg-emerald-900 text-emerald-300';
      default:
        return 'bg-gray-800 text-gray-400';
    }
  };

  // ---- フェーズ: バリデーション中（スピナー表示）----
  if (phase === 'validating') {
    return (
      <div className="flex min-h-screen items-center justify-center">
        <div className="flex flex-col items-center gap-4">
          <Loader2 className="h-10 w-10 animate-spin text-emerald-500" />
          <p className="text-gray-400">トークンを検証中...</p>
        </div>
      </div>
    );
  }

  // ---- フェーズ: エラー ----
  if (phase === 'error') {
    return (
      <div className="flex min-h-screen items-center justify-center px-4">
        <div className="flex w-full max-w-md flex-col items-center gap-4 rounded-2xl border border-red-900 bg-gray-900 p-8">
          <AlertCircle className="h-12 w-12 text-red-500" />
          <h2 className="text-xl font-semibold text-red-400">アクセスエラー</h2>
          <p className="text-center text-sm text-gray-400">{error}</p>
        </div>
      </div>
    );
  }

  // ---- メイン UI ----
  return (
    <div className="flex min-h-screen flex-col bg-gray-950">
      {/* ヘッダー */}
      <header className="border-b border-gray-800 bg-gray-900 px-6 py-4">
        <div className="mx-auto flex max-w-6xl items-center justify-between">
          <div className="flex items-center gap-3">
            <span className="text-xl font-bold text-white">SonaCore Web</span>
            {session && (
              <>
                <span className="text-gray-600">/</span>
                <span className="text-gray-300">{session.name}</span>
              </>
            )}
          </div>
          {/* 録音中の経過時間 */}
          {phase === 'recording' && (
            <div className="flex items-center gap-2 text-red-400">
              <span className="h-2 w-2 animate-pulse rounded-full bg-red-500" />
              <span className="font-mono text-sm">{formatElapsed(elapsedSeconds)}</span>
            </div>
          )}
        </div>
      </header>

      {/* メインコンテンツ（2カラム） */}
      <main className="mx-auto flex w-full max-w-6xl flex-1 gap-6 p-6">
        {/* 左カラム: セッション情報 */}
        <aside className="w-72 flex-shrink-0">
          <div className="rounded-2xl border border-gray-800 bg-gray-900 p-5">
            <h2 className="mb-4 text-sm font-semibold uppercase tracking-wider text-gray-500">
              セッション情報
            </h2>
            {session ? (
              <div className="flex flex-col gap-3">
                <div>
                  <p className="mb-1 text-xs text-gray-500">セッション名</p>
                  <p className="font-medium text-white">{session.name}</p>
                </div>
                <div>
                  <p className="mb-1 text-xs text-gray-500">ステータス</p>
                  <span
                    className={`inline-block rounded-full px-3 py-1 text-xs font-medium ${statusBadgeClass(session.status)}`}
                  >
                    {session.status}
                  </span>
                </div>
                {session.description && (
                  <div>
                    <p className="mb-1 text-xs text-gray-500">説明</p>
                    <p className="text-sm text-gray-300">{session.description}</p>
                  </div>
                )}
                <div>
                  <p className="mb-1 text-xs text-gray-500">セッション ID</p>
                  <p className="break-all font-mono text-xs text-gray-500">{session.id}</p>
                </div>
                <div>
                  <p className="mb-1 text-xs text-gray-500">エントリ数</p>
                  <p className="text-sm text-gray-300">{transcripts.length} 件</p>
                </div>
              </div>
            ) : (
              <p className="text-sm text-gray-500">読み込み中...</p>
            )}
          </div>

          {/* 確定済みトランスクリプト数バッジ */}
          {transcripts.length > 0 && (
            <div className="mt-4 flex items-center gap-2 rounded-xl border border-emerald-900 bg-emerald-950 px-4 py-3">
              <CheckCircle2 className="h-4 w-4 text-emerald-500" />
              <span className="text-sm text-emerald-400">
                {transcripts.length} 件の確定済みテキスト
              </span>
            </div>
          )}
        </aside>

        {/* 右カラム: トランスクリプトストリーム */}
        <div className="flex flex-1 flex-col overflow-hidden rounded-2xl border border-gray-800 bg-gray-900">
          <div className="border-b border-gray-800 px-5 py-3">
            <h2 className="text-sm font-semibold uppercase tracking-wider text-gray-500">
              トランスクリプト
            </h2>
          </div>

          <div className="flex-1 overflow-y-auto p-5">
            {transcripts.length === 0 && !pendingTranscript ? (
              <div className="flex h-full items-center justify-center">
                <p className="text-sm text-gray-600">
                  {phase === 'recording'
                    ? '音声を待機中...'
                    : '録音を開始するとテキストがここに表示されます'}
                </p>
              </div>
            ) : (
              <div className="flex flex-col gap-3">
                {/* 確定済みエントリ */}
                {transcripts.map((entry) => (
                  <div
                    key={entry.id}
                    className="rounded-xl border border-gray-800 bg-gray-800/50 px-4 py-3"
                  >
                    <div className="mb-1 flex items-center gap-2">
                      <span className="text-xs font-semibold text-emerald-400">
                        {entry.speaker}
                      </span>
                      <span className="text-xs text-gray-600">{entry.timestamp}</span>
                    </div>
                    <p className="text-sm leading-relaxed text-gray-200">{entry.text}</p>
                  </div>
                ))}

                {/* 未確定（interim）エントリ */}
                {pendingTranscript && (
                  <div className="rounded-xl border border-yellow-900/50 bg-yellow-950/30 px-4 py-3 opacity-70">
                    <div className="mb-1 flex items-center gap-2">
                      <span className="text-xs font-semibold text-yellow-500">
                        {pendingTranscript.speaker}
                      </span>
                      <span className="text-xs text-gray-600">認識中...</span>
                    </div>
                    <p className="text-sm italic leading-relaxed text-gray-400">
                      {pendingTranscript.text}
                    </p>
                  </div>
                )}

                {/* 自動スクロール用のアンカー */}
                <div ref={transcriptEndRef} />
              </div>
            )}
          </div>
        </div>
      </main>

      {/* フッター: 操作ボタン */}
      <footer className="border-t border-gray-800 bg-gray-900 px-6 py-4">
        <div className="mx-auto flex max-w-6xl items-center justify-between">
          <div className="flex items-center gap-3">
            {/* 録音開始ボタン（ready 状態） */}
            {phase === 'ready' && (
              <button
                onClick={handleStartRecording}
                className="flex items-center gap-2 rounded-xl bg-emerald-600 px-6 py-3 font-semibold text-white transition-colors hover:bg-emerald-500 active:bg-emerald-700"
              >
                <Mic className="h-5 w-5" />
                録音開始
              </button>
            )}

            {/* 録音停止ボタン（recording 状態） */}
            {phase === 'recording' && (
              <button
                onClick={handleStopRecording}
                className="flex animate-pulse items-center gap-2 rounded-xl bg-red-600 px-6 py-3 font-semibold text-white transition-colors hover:bg-red-500 active:bg-red-700"
              >
                <Square className="h-5 w-5 fill-current" />
                録音停止
                <span className="ml-1 font-mono text-sm opacity-80">
                  {formatElapsed(elapsedSeconds)}
                </span>
              </button>
            )}

            {/* 処理中スピナー（finalizing 状態） */}
            {phase === 'finalizing' && (
              <div className="flex items-center gap-2 text-gray-400">
                <Loader2 className="h-5 w-5 animate-spin" />
                <span>保存中...</span>
              </div>
            )}
          </div>

          {/* Excel 出力ボタン（ready かつ completed のとき表示） */}
          {phase === 'ready' && session?.status === 'completed' && (
            <button
              onClick={handleExport}
              className="flex items-center gap-2 rounded-xl border border-gray-700 bg-gray-800 px-5 py-3 text-sm font-medium text-gray-300 transition-colors hover:bg-gray-700 hover:text-white"
            >
              <FileSpreadsheet className="h-4 w-4" />
              Excel 出力
            </button>
          )}
        </div>
      </footer>
    </div>
  );
}
