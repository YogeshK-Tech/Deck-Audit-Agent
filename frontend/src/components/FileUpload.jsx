import React, { useState, useCallback } from 'react';
import { useDropzone } from 'react-dropzone';
import axios from 'axios';
import { Upload, FileText, FileSpreadsheet, CheckCircle, AlertCircle } from 'lucide-react';

const FileUpload = ({ onFilesUploaded }) => {
  const [pptFile, setPptFile] = useState(null);
  const [excelFiles, setExcelFiles] = useState([]);
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState(null);

  const onDropPPT = useCallback((acceptedFiles) => {
    if (acceptedFiles.length > 0) {
      setPptFile(acceptedFiles[0]);
      setError(null);
    }
  }, []);

  const onDropExcel = useCallback((acceptedFiles) => {
    setExcelFiles(prevFiles => [...prevFiles, ...acceptedFiles]);
    setError(null);
  }, []);

  const { getRootProps: getPPTRootProps, getInputProps: getPPTInputProps, isDragActive: isPPTDragActive } = useDropzone({
    onDrop: onDropPPT,
    accept: {
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'],
      'application/vnd.ms-powerpoint': ['.ppt']
    },
    maxFiles: 1
  });

  const { getRootProps: getExcelRootProps, getInputProps: getExcelInputProps, isDragActive: isExcelDragActive } = useDropzone({
    onDrop: onDropExcel,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls'],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.template': ['.xltx']
    },
    multiple: true
  });

  const removeExcelFile = (indexToRemove) => {
    setExcelFiles(files => files.filter((_, index) => index !== indexToRemove));
  };

  const handleUpload = async () => {
    if (!pptFile || excelFiles.length === 0) {
      setError('Please upload both a PowerPoint file and at least one Excel file');
      return;
    }

    setUploading(true);
    setError(null);

    try {
      const formData = new FormData();
      formData.append('ppt_file', pptFile);
      
      excelFiles.forEach((file) => {
        formData.append('excel_files', file);
      });

      const response = await axios.post('/upload-files', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      });

      onFilesUploaded(response.data);
    } catch (err) {
      setError(err.response?.data?.detail || 'Failed to upload files');
    } finally {
      setUploading(false);
    }
  };

  return (
    <div className="file-upload-container">
      <div className="upload-section">
        <h2>📥 Upload Files</h2>
        <p>Upload your PowerPoint presentation and Excel data files to begin the audit</p>

        {/* PowerPoint Upload */}
        <div className="upload-box">
          <h3><FileText className="icon" /> PowerPoint File</h3>
          <div 
            {...getPPTRootProps()} 
            className={`dropzone ${isPPTDragActive ? 'active' : ''} ${pptFile ? 'has-file' : ''}`}
          >
            <input {...getPPTInputProps()} />
            {pptFile ? (
              <div className="file-info">
                <CheckCircle className="icon success" />
                <span>{pptFile.name}</span>
                <button onClick={(e) => { e.stopPropagation(); setPptFile(null); }}>Remove</button>
              </div>
            ) : (
              <div className="upload-prompt">
                <Upload className="icon" />
                <p>{isPPTDragActive ? 'Drop the file here' : 'Drag & drop a PowerPoint file here, or click to select'}</p>
                <p className="file-types">Supports: .pptx, .ppt</p>
              </div>
            )}
          </div>
        </div>

        {/* Excel Upload */}
        <div className="upload-box">
          <h3><FileSpreadsheet className="icon" /> Excel Files</h3>
          <div 
            {...getExcelRootProps()} 
            className={`dropzone ${isExcelDragActive ? 'active' : ''} ${excelFiles.length > 0 ? 'has-file' : ''}`}
          >
            <input {...getExcelInputProps()} />
            <div className="upload-prompt">
              <Upload className="icon" />
              <p>{isExcelDragActive ? 'Drop the files here' : 'Drag & drop Excel files here, or click to select'}</p>
              <p className="file-types">Supports: .xlsx, .xls • Multiple files allowed</p>
            </div>
          </div>

          {excelFiles.length > 0 && (
            <div className="file-list">
              <h4>Uploaded Excel Files:</h4>
              {excelFiles.map((file, index) => (
                <div key={index} className="file-item">
                  <FileSpreadsheet className="icon" />
                  <span>{file.name}</span>
                  <button onClick={() => removeExcelFile(index)}>Remove</button>
                </div>
              ))}
            </div>
          )}
        </div>

        {error && (
          <div className="error-message">
            <AlertCircle className="icon" />
            {error}
          </div>
        )}

        <button 
          className="upload-button"
          onClick={handleUpload}
          disabled={!pptFile || excelFiles.length === 0 || uploading}
        >
          {uploading ? 'Processing Files...' : 'Start Audit'}
        </button>
      </div>
    </div>
  );
};

export default FileUpload;