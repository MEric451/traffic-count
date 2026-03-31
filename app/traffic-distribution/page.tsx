'use client';

import { useState } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Button } from '@/components/ui/button';
import { Eye, Copy, FileUp, Settings, Trash2 } from 'lucide-react';
import { Label } from '@/components/ui/label';
import { toast } from 'sonner';
import { Toaster } from 'sonner';
import LaneDistributionInput from '@/components/LaneDistributionInput';
import PreviewModal from '@/components/PreviewModal';
import ProcessingModal from '@/components/ProcessingModal';
import DeleteConfirmationDialog from '@/components/DeleteConfirmationDialog';
import FileProcessingLoader from '@/components/FileProcessingLoader';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from '@/components/ui/dialog';

export default function TrafficDistributionPage() {
  const [file, setFile] = useState<File | null>(null);
  const [scriptType, setScriptType] = useState<'12hour' | '16hour' | '24hour'>('12hour');
  const [outputFilename, setOutputFilename] = useState<string>('');
  const [lanes, setLanes] = useState<string[]>([]);
  const [laneDistributions, setLaneDistributions] = useState<Record<string, number[]>>({});
  const [showPreview, setShowPreview] = useState(false);
  const [processingStatus, setProcessingStatus] = useState<'idle' | 'processing' | 'success' | 'error'>('idle');
  const [error, setError] = useState<string>('');
  const [processedFileBlob, setProcessedFileBlob] = useState<Blob | null>(null);
  const [isDragOver, setIsDragOver] = useState(false);
  const [isProcessingFile, setIsProcessingFile] = useState(false);
  const [showDeleteDialog, setShowDeleteDialog] = useState(false);

  // Always 12 vehicle classes regardless of hour type
  const expectedCount = 12;

  const processFile = async (uploadedFile: File) => {
    setFile(uploadedFile);
    setLanes([]);
    setLaneDistributions({});
    setIsProcessingFile(true);
    
    // Auto-fill output filename based on uploaded file name (without extension)
    const fileNameWithoutExt = uploadedFile.name.replace(/\.xlsx?$/i, '');
    setOutputFilename(`${fileNameWithoutExt} - Modified`);

    // Detect lanes from uploaded file
    const formData = new FormData();
    formData.append('file', uploadedFile);

    try {
      const response = await fetch('/api/detect-lanes', {
        method: 'POST',
        body: formData,
      });

      const result = await response.json();

      if (result.success) {
        if (result.lanes && result.lanes.length > 0) {
          // Small delay for smooth UX
          await new Promise(resolve => setTimeout(resolve, 600));
          
          setLanes(result.lanes);
          // Initialize empty distributions
          const initialDist: Record<string, number[]> = {};
          result.lanes.forEach((lane: string) => {
            initialDist[lane] = [];
          });
          setLaneDistributions(initialDist);
        } else {
          toast.error('No lanes/bounds found in the uploaded file. Please check your Excel file format.');
        }
      } else {
        toast.error(`Error detecting lanes: ${result.error}`);
      }
    } catch (err) {
      toast.error('Failed to detect lanes from file');
      console.error(err);
    } finally {
      setIsProcessingFile(false);
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFile = e.target.files?.[0];
    if (!uploadedFile) return;
    await processFile(uploadedFile);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    
    const droppedFile = e.dataTransfer.files[0];
    if (droppedFile && droppedFile.name.endsWith('.xlsx')) {
      await processFile(droppedFile);
    } else {
      toast.error('Please drop a valid excel  file');
    }
  };

  const handleDistributionChange = (lane: string, distribution: number[]) => {
    setLaneDistributions((prev) => ({
      ...prev,
      [lane]: distribution,
    }));
  };

  const handlePreview = () => {
    // Validate all lanes have distributions
    const missingLanes = lanes.filter((lane) => !laneDistributions[lane] || laneDistributions[lane].length === 0);
    
    if (missingLanes.length > 0) {
      toast.error(`Please set distribution for: ${missingLanes.join(', ')}`);
      return;
    }

    if (!outputFilename.trim()) {
      toast.error('Please enter an output filename');
      return;
    }

    setShowPreview(true);
  };

  const handleProcess = async () => {
    setShowPreview(false);
    setProcessingStatus('processing');
    setError('');

    const formData = new FormData();
    formData.append('file', file!);
    formData.append('scriptType', scriptType);
    formData.append('outputFilename', outputFilename.endsWith('.xlsx') ? outputFilename : `${outputFilename}.xlsx`);
    formData.append('totals', JSON.stringify(laneDistributions));

    try {
      const response = await fetch('/api/process-traffic', {
        method: 'POST',
        body: formData,
      });

      if (response.ok) {
        const blob = await response.blob();
        setProcessedFileBlob(blob);
        setProcessingStatus('success');
      } else {
        const errorData = await response.json();
        setError(errorData.error || 'Processing failed');
        setProcessingStatus('error');
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Unknown error occurred');
      setProcessingStatus('error');
    }
  };

  const handleDownload = () => {
    if (!processedFileBlob) return;

    const url = URL.createObjectURL(processedFileBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = outputFilename.endsWith('.xlsx') ? outputFilename : `${outputFilename}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleCloseProcessing = () => {
    setProcessingStatus('idle');
    setProcessedFileBlob(null);
  };

  const handleDeleteFile = () => {
    // Reset to initial state
    setFile(null);
    setOutputFilename('');
    setLanes([]);
    setLaneDistributions({});
    setProcessingStatus('idle');
    setError('');
    setProcessedFileBlob(null);
    setShowPreview(false);
    setShowDeleteDialog(false);
    
    // Reset file input
    const fileInput = document.getElementById('file-upload') as HTMLInputElement;
    if (fileInput) {
      fileInput.value = '';
    }
    
    toast.success('File removed successfully');
  };

  const handleDeleteClick = () => {
    setShowDeleteDialog(true);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-purple-50 py-4 sm:py-8">
      <div className="container mx-auto px-4 max-w-5xl">
        {/* Header with gradient */}
        <div className="mb-8 text-center">
          <div className="inline-flex items-center gap-3 mb-3">
            <div className="w-12 h-12 flex items-center justify-center text-white text-2xl shadow-lg">
              🚗
            </div>
            <h1 className="text-3xl sm:text-4xl font-bold bg-gradient-to-r from-blue-600 to-purple-600 bg-clip-text text-transparent">
              Traffic Distribution Generator
            </h1>
          </div>
          <p className="text-gray-600 max-w-2xl mx-auto">
            Generate realistic traffic count distributions for your Excel files with our intelligent processing engine
          </p>
        </div>

        {/* Step 1 & 2: Upload File and Processing Options */}
        <Card className="mb-6 shadow-xl border-0 bg-white/80 backdrop-blur-sm">
          <CardContent className="p-6">
            <div className="flex flex-col lg:flex-row gap-8">
              {/* Upload Section */}
              <div className="flex-1">
                <h3 className="text-lg font-semibold mb-4 text-gray-800">
                  <FileUp className="inline w-5 h-5 mb-2" /> Upload Excel File</h3>
                <div 
                  className={`border-2 border-dashed rounded-xl p-8 text-center transition-all duration-300 transform ${
                    isDragOver 
                      ? 'border-blue-500 bg-gradient-to-br from-blue-50 to-purple-50 shadow-lg scale-[1.02]' 
                      : 'border-gray-300 hover:border-blue-400 hover:bg-gray-50'
                  }`}
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onDrop={handleDrop}
                >
                  <div className={`w-16 h-16 mx-auto mb-4 rounded-full flex items-center justify-center transition-all duration-300 ${
                    isDragOver ? 'bg-blue-500 text-white' : 'bg-gray-100 text-gray-400'
                  }`}>
                    <FileUp className="w-8 h-8" />
                  </div>
                  <p className="mb-4 text-gray-600 font-medium">
                    {isDragOver ? 'Drop your file here!' : 'Drag & drop your Excel file here, or'}
                  </p>
                  <input
                    type="file"
                    accept=".xlsx"
                    onChange={handleFileUpload}
                    className="hidden"
                    id="file-upload"
                  />
                  <label htmlFor="file-upload" className="cursor-pointer">
                    <Button className="bg-gradient-to-r hover:scale-[1.02] from-blue-500 to-purple-600 hover:from-blue-600 hover:to-purple-700 shadow-lg" asChild>
                      <span>Browse Files</span>
                    </Button>
                  </label>
                  {file && !isProcessingFile && (
                    <div className="mt-4 p-3 bg-green-50 rounded-lg border border-green-200">
                      <div className="flex items-start gap-2 text-sm text-green-700">
                        <div className="w-5 h-5 flex items-center justify-center text-green-500 text-xs">✓</div>
                        <div className="flex-1">
                          <span className="font-medium">Uploaded: {file.name}</span>
                        </div>
                        <button
                          onClick={handleDeleteClick}
                          className="p-1 hover:bg-red-200 rounded transition-colors text-red-600 hover:text-red-700"
                          title="Remove file"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </div>
                  )}
                  
                  {/* File Processing Loader */}
                  {isProcessingFile && (
                    <FileProcessingLoader fileName={file?.name} />
                  )}
                  {lanes.length > 0 && (
                    <div className="mt-3 p-3 bg-blue-50 rounded-lg border border-blue-200">
                      <div className="flex items-center gap-2 text-sm text-blue-700 font-medium">
                        <div className="w-5 h-5 flex items-center justify-center text-blue-500 text-xs">✓</div>
                        <span>Detected lanes: {lanes.join(', ')}</span>
                      </div>
                    </div>
                  )}
                </div>
              </div>

              {/* Processing Options Section */}
              {file && lanes.length > 0 && (
                <div className="flex-1 animate-in slide-in-from-right duration-500">
                  <h3 className="text-lg font-semibold mb-4 text-gray-800">
                    <Settings className="inline w-5 h-5 mb-2 mr-2" />
                    Processing Options</h3>
                  <div className="space-y-6">
                    <div>
                      <label className="text-sm font-medium mb-3 block text-gray-900">Script Type</label>
                      <RadioGroup value={scriptType} onValueChange={(value) => setScriptType(value as '12hour' | '16hour' | '24hour')} className="space-y-2">
                        <div className="flex items-center space-x-2 p-3 transition-colors">
                          <RadioGroupItem value="12hour" id="12hour"/>
                          <Label htmlFor="12hour" className="flex-1 cursor-pointer text-gray-800">12-Hour (7AM-7PM)</Label>
                        </div>
                        <div className="flex items-center space-x-2 p-3 transition-colors">
                          <RadioGroupItem value="16hour" id="16hour"/>
                          <Label htmlFor="16hour" className="flex-1 cursor-pointer text-gray-800">16-Hour (6AM-10PM)</Label>
                        </div>
                        <div className="flex items-center space-x-2 p-3 transition-colors">
                          <RadioGroupItem value="24hour" id="24hour"  />
                          <Label htmlFor="24hour" className="flex-1 cursor-pointer text-gray-800">24-Hour (Full Day)</Label>
                        </div>
                      </RadioGroup>
                    </div>

                    <div>
                      <label className="text-sm font-medium mb-3 block text-gray-700">Output Filename</label>
                      <Input
                        placeholder="Enter your desired output filename"
                        value={outputFilename}
                        onChange={(e) => setOutputFilename(e.target.value)}
                        className="border-gray-300 transition-colors focus:border-blue-500 focus:ring-2 focus:ring-blue-200"
                      />
                    </div>
                  </div>
                </div>
              )}
            </div>
          </CardContent>
        </Card>

        {/* Step 2: Traffic Distribution */}
        {file && lanes.length > 0 && (
          <Card className="mb-6 shadow-xl border-0 bg-white/80 backdrop-blur-sm animate-in slide-in-from-bottom duration-500">
            <CardHeader className="bg-gradient-to-r from-purple-500 to-pink-600 text-white rounded-t-lg p-3">
              <CardTitle>
                Traffic Distribution For All Bounds
              </CardTitle>
            </CardHeader>
            <CardContent className="p-3">
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                {lanes.map((lane, index) => (
                  <div key={lane} className="animate-in slide-in-from-left duration-500" style={{animationDelay: `${index * 100}ms`}}>
                    <LaneDistributionInput
                      laneName={lane}
                      expectedCount={expectedCount}
                      onDistributionChange={(dist) => handleDistributionChange(lane, dist)}
                    />
                  </div>
                ))}
              </div>
            </CardContent>
          </Card>
        )}

        {/* Action Buttons */}
        {file && lanes.length > 0 && (
          <div className="flex flex-col sm:flex-row gap-4 sm:justify-center animate-in slide-in-from-bottom duration-500">
            <Button 
         
              onClick={handlePreview} 
              className="gap-2 w-full sm:w-auto border-gray-800 bg-gray-200 text-black hover:bg-gray-300 transform transition-all hover:scale-105 duration-200"
            >
              <Eye className="w-4 h-4" />
              Preview Summary
            </Button>
            <Button 
              onClick={handlePreview} 
              size="lg" 
              className="gap-2 w-full sm:w-auto bg-gradient-to-r from-green-500 to-blue-600 hover:from-green-600 hover:to-blue-700 shadow-lg transform hover:scale-105 transition-all duration-200"
            >
              Generate Excel File
            </Button>
          </div>
        )}

        {/* Modals */}
        <PreviewModal
          open={showPreview}
          onClose={() => setShowPreview(false)}
          onConfirm={handleProcess}
          scriptType={scriptType}
          outputFilename={outputFilename}
          totals={laneDistributions}
        />

        <ProcessingModal
          open={processingStatus !== 'idle'}
          status={processingStatus === 'processing' ? 'processing' : processingStatus === 'success' ? 'success' : 'error'}
          onClose={handleCloseProcessing}
          onDownload={handleDownload}
          error={error}
        />

        <DeleteConfirmationDialog
          open={showDeleteDialog}
          onClose={() => setShowDeleteDialog(false)}
          onConfirm={handleDeleteFile}
          fileName={file?.name}
        />
      </div>
      <Toaster position="top-right" />
    </div>
  );
}
