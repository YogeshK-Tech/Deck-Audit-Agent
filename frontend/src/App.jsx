import React, { useState } from 'react';
import FileUpload from './components/FileUpload';
import AuditDashboard from './components/AuditDashboard';
import './App.css';

function App() {
  const [currentStep, setCurrentStep] = useState('upload'); // 'upload', 'audit', 'results'
  const [auditData, setAuditData] = useState(null);
  const [sessionId, setSessionId] = useState(null);

  const handleFilesUploaded = (data) => {
    setSessionId(data.session_id);
    setCurrentStep('audit');
  };

  const handleAuditComplete = (results) => {
    setAuditData(results);
    setCurrentStep('results');
  };

  const resetApp = () => {
    setCurrentStep('upload');
    setAuditData(null);
    setSessionId(null);
  };

  return (
    <div className="App">
      <header className="app-header">
        <h1>🔧 Deck Auditor</h1>
        <p>Intelligent AI Agent for PowerPoint Number Verification</p>
      </header>

      <main className="app-main">
        {currentStep === 'upload' && (
          <FileUpload onFilesUploaded={handleFilesUploaded} />
        )}
        
        {(currentStep === 'audit' || currentStep === 'results') && (
          <AuditDashboard 
            sessionId={sessionId}
            auditData={auditData}
            onAuditComplete={handleAuditComplete}
            onReset={resetApp}
          />
        )}
      </main>

      <footer className="app-footer">
        <p>Built with React + FastAPI • Powered by AI</p>
      </footer>
    </div>
  );
}

export default App;