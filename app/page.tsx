// トップページ: /record/[sessionId] への案内メッセージを表示する
export default function HomePage() {
  return (
    <div className="flex min-h-screen flex-col items-center justify-center gap-6 px-4">
      <div className="text-center">
        <h1 className="mb-2 text-4xl font-bold tracking-tight text-white">
          SonaCore Web
        </h1>
        <p className="text-lg text-gray-400">ブラウザ録音・文字起こしアプリ</p>
      </div>

      <div className="w-full max-w-md rounded-2xl border border-gray-800 bg-gray-900 p-6">
        <h2 className="mb-3 text-lg font-semibold text-gray-200">
          このページの使い方
        </h2>
        <p className="mb-4 text-sm text-gray-400 leading-relaxed">
          録音を開始するには、SonaCore デスクトップアプリから発行されたリンクを使用してください。
          URLの形式は以下の通りです：
        </p>
        <code className="block rounded-lg bg-gray-800 px-4 py-3 text-sm text-emerald-400 break-all">
          /record/[セッションID]?token=[起動トークン]
        </code>
        <p className="mt-4 text-xs text-gray-500">
          トークンは一度のみ使用可能です。セッションが有効な間はリロードしても継続できます。
        </p>
      </div>
    </div>
  );
}
