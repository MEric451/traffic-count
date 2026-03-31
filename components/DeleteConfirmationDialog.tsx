import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Trash2 } from 'lucide-react';

interface DeleteConfirmationDialogProps {
  open: boolean;
  onClose: () => void;
  onConfirm: () => void;
  fileName?: string;
}

export default function DeleteConfirmationDialog({ 
  open, 
  onClose, 
  onConfirm, 
  fileName 
}: DeleteConfirmationDialogProps) {
  return (
    <Dialog open={open} onOpenChange={onClose}>
      <DialogContent className="sm:max-w-md">
        <DialogHeader>
          <DialogTitle className="flex items-center gap-2 text-red-600">
            <Trash2 className="w-5 h-5" />
            Remove File
          </DialogTitle>
          <DialogDescription>
            Are you sure you want to remove <span className="font-medium">
              {fileName ? `"${fileName}"` : 'the uploaded file'}?
            </span> This will reset all your progress and you&apos;ll need to start over.
          </DialogDescription>
        </DialogHeader>
        <DialogFooter className="flex gap-3 sm:gap-2">
          <Button
            variant="outline"
            onClick={onClose}
            className="flex-1 sm:flex-none"
          >
            Cancel
          </Button>
          <Button
            variant="destructive"
            onClick={onConfirm}
            className="flex-1 sm:flex-none bg-red-600 hover:bg-red-700"
          >
            Remove File
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}