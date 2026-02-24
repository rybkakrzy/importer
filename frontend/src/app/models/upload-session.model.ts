export interface UploadSession {
  id: string;
  fileName: string;
  status: 'Processing' | 'Completed' | 'Failed';
  startedAt: string;
  completedAt?: string;
  filesCount?: number;
  totalSize?: number;
  errorMessage?: string;
}
