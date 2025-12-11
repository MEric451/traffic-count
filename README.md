## Traffic Count Modifier
A Next.js application for uploading, analyzing, and modifying Excel traffic count data.

**Features**
•	Excel File Processing – Upload .xlsx files and automatically modify traffic count values
•	Dynamic Bound Detection – Detects and processes any bound/location name
•	Percentage-based Adjustments – Increase or decrease counts by a custom percentage
•	Multi-Sheet Support – Processes all time-based sheets (7-8AM through 6-7PM)
•	Drag & Drop Upload – Smooth drag-and-drop interface for file submission
•	Dark Mode UI – Switch between light and dark themes
•	Real-time Processing Logs – View detailed Python logs after each operation
•	Format Preservation – Maintains original formatting, formulas, and colors
•	Responsive Design – Fully optimized for desktop and mobile
•	Python Backend Integration – Uses openpyxl for accurate Excel manipulation
  
**Tech Stack**
  1.	Framework: Next.js 15 / App Router
  2.	Language: TypeScript
  3.	Styling: Tailwind CSS
  4.	Icons: Lucide React
  5.	Backend Processing: Python (openpyxl)
  6.	File Handling: Node.js File System API
  7.	Runtime: Node.js server functions (for Python execution)

**Getting Started**
  1.	Install dependencies:
  npm install
  2.	Ensure python and packages are installed:
  pip install openpyxl
  2.	Start development server:
  npm run dev
  3.	Open http://localhost:3000 in your browser.

**Project Structure**
  app/
  ├── excel-modifier/        # Main UI for uploading and processing files
  ├── api/
  │   └── modify-excel/      # API route that executes the Python script
  ├── modify_excel.py        # Python script responsible for Excel processing
  public/                    # Static assets

**Supported Excel Format**
  The application expects:
    •	Time-based sheet names
    •	7-8AM
    •	8-9AM
    •	9-10AM
    •	...
    •	6-7PM
    •	Bound/Location names in column A
    •	Traffic counts in columns B–M
    •	Original formatting and formulas are preserved


## Deployment
Deploy to Vercel by importing the Github repository.
Vercel auto-detects Next.js and provides a live URL.
