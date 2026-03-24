'use client';

import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Download, Loader2, CheckCircle2, XCircle } from 'lucide-react';

interface ProcessingModalProps {
  open: boolean;
  status: 'processing' | 'success' | 'error';
  onClose: () => void;
  onDownload?: () => void;
  error?: string;
}

export default function ProcessingModal({
  open,
  status,
  onClose,
  onDownload,
  error,
}: ProcessingModalProps) {
  return (
    <Dialog open={open} onOpenChange={onClose}>
      <DialogContent className="max-w-md">
        <DialogHeader>
          <DialogTitle>
            {status === 'processing' && 'Generating Excel File...'}
            {status === 'success' && 'Success!'}
            {status === 'error' && 'Error'}
          </DialogTitle>
          <DialogDescription>
            {status === 'processing' && 'Please wait while we process your traffic distribution'}
            {status === 'success' && 'Your Excel file has been generated successfully'}
            {status === 'error' && 'An error occurred while processing'}
          </DialogDescription>
        </DialogHeader>

        <div className="py-6 flex flex-col items-center justify-center">
          {status === 'processing' && (
            <Loader2 className="w-16 h-16 animate-spin text-blue-500" />
          )}
          {status === 'success' && (
            <CheckCircle2 className="w-16 h-16 text-green-500" />
          )}
          {status === 'error' && (
            <XCircle className="w-16 h-16 text-red-500" />
          )}

          {status === 'processing' && (
            <p className="mt-4 text-sm text-gray-600">
              Processing traffic distribution...
            </p>
          )}

          {status === 'success' && (
            <div className="mt-4 text-center">
              <p className="text-sm text-gray-600 mb-4">
                Your file is ready for download
              </p>
              <Button onClick={onDownload} className="gap-2">
                <Download className="w-4 h-4" />
                Download File
              </Button>
            </div>
          )}

          {status === 'error' && error && (
            <div className="mt-4 text-center">
              <p className="text-sm text-red-600 bg-red-50 p-3 rounded-md">
                {error}
              </p>
              <Button onClick={onClose} variant="outline" className="mt-4">
                Close
              </Button>
            </div>
          )}
        </div>
      </DialogContent>
    </Dialog>
  );
}
