@echo off
REM Batch script to manage document vectorization
REM Save this as "process_documents.bat"

:menu
echo.
echo Document Vector Database Manager
echo ================================
echo 1. Process new documents
echo 2. Start continuous monitoring
echo 3. Search documents
echo 4. Setup environment
echo 5. Exit
echo.
set /p choice="Select an option (1-5): "

if "%choice%"=="1" goto process
if "%choice%"=="2" goto monitor
if "%choice%"=="3" goto search
if "%choice%"=="4" goto setup
if "%choice%"=="5" goto exit
goto menu

:process
echo.
echo Processing new documents...
python -c "from document_vectorizer import DocumentVectorizer; v = DocumentVectorizer('./batch_documents', './vector_database'); stats = v.batch_process_folder(); print(f'Processed: {stats}')"
pause
goto menu

:monitor
echo.
echo Starting continuous monitoring...
echo Press Ctrl+C to stop monitoring
python -c "from document_vectorizer import DocumentVectorizer; v = DocumentVectorizer('./batch_documents', './vector_database'); v.monitor_folder(check_interval=60)"
pause
goto menu

:search
echo.
set /p query="Enter search query: "
python -c "from document_vectorizer import DocumentVectorizer; v = DocumentVectorizer('./batch_documents', './vector_database'); results = v.search_documents('%query%'); [print(f'\n{i+1}. {r[\"metadata\"][\"filename\"]}: {r[\"content\"][:100]}...') for i, r in enumerate(results)]"
pause
goto menu

:setup
echo.
echo Setting up environment...
echo Creating required directories...
mkdir batch_documents 2>nul
mkdir vector_database 2>nul

echo.
echo Installing required Python packages...
pip install PyPDF2 python-docx openpyxl python-pptx pandas
pip install langchain chromadb faiss-cpu
pip install sentence-transformers openai

echo.
echo Setup complete!
pause
goto menu

:exit
exit