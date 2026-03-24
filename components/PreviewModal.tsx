'use client';

import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { VEHICLE_CLASSES } from '@/lib/utils';

interface PreviewModalProps {
  open: boolean;
  onClose: () => void;
  onConfirm: () => void;
  scriptType: string;
  outputFilename: string;
  totals: Record<string, number[]>;
}

export default function PreviewModal({
  open,
  onClose,
  onConfirm,
  scriptType,
  outputFilename,
  totals,
}: PreviewModalProps) {
  const grandTotal = Object.values(totals).reduce(
    (sum, dist) => sum + dist.reduce((a, b) => a + b, 0),
    0
  );

  const classTotals = VEHICLE_CLASSES.map((_, index) =>
    Object.values(totals).reduce((sum, dist) => sum + (dist[index] || 0), 0)
  );

  return (
    <Dialog open={open} onOpenChange={onClose}>
      <DialogContent className="max-w-[95vw] sm:max-w-3xl max-h-[90vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle className="text-lg sm:text-xl">Review Before Processing</DialogTitle>
          <DialogDescription className="text-sm">
            Please review the configuration before generating the Excel file
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-4">
          {/* Configuration */}
          <div className="border rounded-lg p-3 sm:p-4 space-y-2 text-sm">
            <div className="flex flex-col sm:flex-row sm:justify-between gap-1">
              <span className="font-medium">Script Type:</span>
              <span>{scriptType}</span>
            </div>
            <div className="flex flex-col sm:flex-row sm:justify-between gap-1">
              <span className="font-medium">Output File:</span>
              <span className="text-xs sm:text-sm break-all">{outputFilename}</span>
            </div>
          </div>

          {/* Lane Distributions */}
          <div className="border rounded-lg p-3 sm:p-4">
            <h3 className="font-medium mb-3 text-sm sm:text-base">Lane Distributions</h3>
            <div className="space-y-2">
              {Object.entries(totals).map(([lane, dist]) => (
                <div key={lane} className="flex flex-col sm:flex-row sm:justify-between sm:items-center py-2 border-b last:border-b-0 gap-1">
                  <span className="font-medium text-sm">{lane}:</span>
                  <span className="text-base sm:text-lg">{dist.reduce((a, b) => a + b, 0).toLocaleString()} vehicles</span>
                </div>
              ))}
              <div className="flex flex-col sm:flex-row sm:justify-between sm:items-center py-2 font-bold text-base sm:text-lg border-t-2 gap-1">
                <span>TOTAL:</span>
                <span>{grandTotal.toLocaleString()} vehicles</span>
              </div>
            </div>
          </div>

          {/* Distribution Breakdown */}
          <div className="border rounded-lg p-3 sm:p-4">
            <h3 className="font-medium mb-3 text-sm sm:text-base">Distribution Breakdown</h3>
            <div className="overflow-x-auto -mx-3 sm:mx-0">
              <table className="w-full text-xs sm:text-sm min-w-[500px]">
                <thead>
                  <tr className="border-b">
                    <th className="text-left py-2 px-2">Vehicle Class</th>
                    {Object.keys(totals).map((lane) => (
                      <th key={lane} className="text-right py-2 px-1 sm:px-2">{lane.split(' ').pop()}</th>
                    ))}
                    <th className="text-right py-2 px-1 sm:px-2 font-bold">Total</th>
                  </tr>
                </thead>
                <tbody>
                  {VEHICLE_CLASSES.map((className, index) => (
                    <tr key={index} className="border-b">
                      <td className="py-2 px-2 text-xs">{className}</td>
                      {Object.values(totals).map((dist, laneIndex) => (
                        <td key={laneIndex} className="text-right py-2 px-1 sm:px-2">
                          {dist[index]?.toLocaleString() || 0}
                        </td>
                      ))}
                      <td className="text-right py-2 px-1 sm:px-2 font-medium">
                        {classTotals[index].toLocaleString()}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <DialogFooter className="flex-col sm:flex-row gap-2">
          <Button variant="outline" onClick={onClose} className="w-full sm:w-auto">
            Cancel
          </Button>
          <Button onClick={onConfirm} className="w-full sm:w-auto">
            Confirm & Generate
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
