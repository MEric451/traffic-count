interface FileProcessingLoaderProps {
  fileName?: string;
}

export default function FileProcessingLoader({ fileName }: FileProcessingLoaderProps) {
  return (
    <div className="mt-4 p-6 bg-gradient-to-r from-blue-50 to-purple-50 rounded-lg border border-blue-200">
      <div className="flex flex-col items-center space-y-4">
        <div className="relative">
          <div className="w-12 h-12 border-4 border-blue-200 rounded-full animate-spin">
            <div className="absolute top-0 left-0 w-12 h-12 border-4 border-transparent border-t-blue-500 rounded-full animate-spin"></div>
          </div>
          <div className="absolute inset-0 flex items-center justify-center">
            <span className="text-lg animate-pulse">🚗</span>
          </div>
        </div>
        <div className="text-center">
          <p className="text-sm font-medium text-blue-700 animate-pulse">
            Analyzing Excel structure...
          </p>
          <p className="text-xs text-blue-600 mt-1">
            Detecting lanes and bounds
          </p>
          {fileName && (
            <p className="text-xs text-gray-500 mt-1">
              Processing: {fileName}
            </p>
          )}
        </div>
        <div className="flex space-x-1">
          <div className="w-2 h-2 bg-blue-500 rounded-full animate-bounce" style={{animationDelay: '0ms'}}></div>
          <div className="w-2 h-2 bg-blue-500 rounded-full animate-bounce" style={{animationDelay: '150ms'}}></div>
          <div className="w-2 h-2 bg-blue-500 rounded-full animate-bounce" style={{animationDelay: '300ms'}}></div>
        </div>
      </div>
    </div>
  );
}