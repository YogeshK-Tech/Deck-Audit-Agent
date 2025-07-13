import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { Play, Download, RotateCcw, CheckCircle, XCircle, AlertTriangle, FileText } from 'lucide-react';
import AuditResults from './AuditResults';

const AuditDashboard = ({ sessionId, auditData, onAuditComplete, onReset }) => {
  const [isRunningAudit, setIsRunningAudit] = useState(false);
  const [auditStatus, setAuditStatus] = useState('ready'); // 'ready', 'running', 'completed'
  const [error, setError] = useState(null);

  const runAudit = async () => {
    setIsRunningAudit(true);
    setAuditStatus('running');
    setError(null);

    try {
      const response = await axios.post('/audit');
      setAuditStatus('completed');
      onAuditComplete(response.data);
    } catch (err) {
      setError(err.response?.data?.detail || 'Failed to run audit');
      setAuditStatus('ready');
    } finally {
      setIsRunningAudit(false);
    }
  };

  const downloadReport = async (format) => {
    try {
      const response = await axios.get(`/download-report/${format}`, {
        responseType: 'blob'
      });

      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `audit_report_${sessionId}.${format}`);
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      setError(`Failed to download ${format.toUpperCase()} report`);
    }
  };

  const getStatusIcon = () => {
    switch (auditStatus) {
      case 'running':
        return <div className="loading-spinner" />;
      case 'completed':
        return <CheckCircle className="icon success" />;
      default:
        return <Play className="icon" />;
    }
  };

  const getStatusMessage = () => {
    switch (auditStatus) {
      case 'running':
        return 'Running audit analysis...';
      case 'completed':
        return 'Audit completed successfully';
      default:
        return 'Ready to start audit';
    }
  };

  return (
    <div className="audit-dashboard">
      <div className="dashboard-header">
        <h2>🔍 Audit Dashboard</h2>
        <button className="reset-button" onClick={onReset}>
          <RotateCcw className="icon" />
          New Audit
        </button>
      </div>

      {auditStatus !== 'completed' && (
        <div className="audit-controls">
          <div className="status-section">
            <div className="status-indicator">
              {getStatusIcon()}
              <span>{getStatusMessage()}</span>
            </div>
            
            {error && (
              <div className="error-message">
                <AlertTriangle className="icon" />
                {error}
              </div>
            )}
          </div>

          <button
            className="audit-button"
            onClick={runAudit}
            disabled={isRunningAudit || auditStatus === 'completed'}
          >
            {isRunningAudit ? 'Analyzing...' : 'Run Audit'}
          </button>
        </div>
      )}

      {auditStatus === 'completed' && auditData && (
        <>
          <div className="audit-summary">
            <h3>📊 Audit Summary</h3>
            <div className="summary-stats">
              <div className="stat-item">
                <CheckCircle className="icon success" />
                <div>
                  <span className="stat-number">{auditData.matches}</span>
                  <span className="stat-label">Matches</span>
                </div>
              </div>
              <div className="stat-item">
                <XCircle className="icon error" />
                <div>
                  <span className="stat-number">{auditData.mismatches}</span>
                  <span className="stat-label">Mismatches</span>
                </div>
              </div>
              <div className="stat-item">
                <AlertTriangle className="icon warning" />
                <div>
                  <span className="stat-number">{auditData.untraceable}</span>
                  <span className="stat-label">Untraceable</span>
                </div>
              </div>
              <div className="stat-item">
                <FileText className="icon" />
                <div>
                  <span className="stat-number">{auditData.total_numbers_found}</span>
                  <span className="stat-label">Total Numbers</span>
                </div>
              </div>
            </div>
          </div>

          <div className="download-section">
            <h3>📄 Download Reports</h3>
            <div className="download-buttons">
              <button onClick={() => downloadReport('pdf')} className="download-btn pdf">
                <Download className="icon" />
                Download PDF Report
              </button>
              <button onClick={() => downloadReport('csv')} className="download-btn csv">
                <Download className="icon" />
                Download CSV Report
              </button>
            </div>
          </div>

          <AuditResults results={auditData.audit_results} />
        </>
      )}
    </div>
  );
};

export default AuditDashboard;