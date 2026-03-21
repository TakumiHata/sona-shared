// サーバーコンポーネント: URLパラメータを受け取りクライアントコンポーネントに渡す
import RecordingConsole from './RecordingConsole';

export default async function RecordPage({
  params,
  searchParams,
}: {
  params: Promise<{ sessionId: string }>;
  searchParams: Promise<{ token?: string }>;
}) {
  const { sessionId } = await params;
  const { token } = await searchParams;
  return <RecordingConsole sessionId={sessionId} token={token ?? ''} />;
}
