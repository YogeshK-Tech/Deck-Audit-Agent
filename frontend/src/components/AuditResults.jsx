import React, { useState, useMemo } from 'react';
import { CheckCircle, XCircle, AlertTriangle, Search, Filter } from 'lucide-react';

const AuditResults = ({ results }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;

  const filteredResults = useMemo(() => {
    let filtered = results;

    // Apply search filter
    if (searchTerm) {
      filtered = filtered.filter(result =>
        result.text.toLowerCase().includes(searchTerm.toLowerCase()) ||
        result.context.toLowerCase().includes(searchTerm.toLowerCase()) ||
        (result.excel_sheet && result.excel_sheet.toLowerCase().includes(searchTerm.toLowerCase()))
      );
    }

    // Apply status filter
    if (statusFilter !== 'all') {
      filtered = filtered.filter(result => result.status === statusFilter);
    }

    return filtered;
  }, [results, searchTerm, statusFilter]);

  const paginatedResults = useMemo(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    return filteredResults.slice(startIndex, startIndex + itemsPerPage);
  }, [filteredResults, currentPage]);

  const totalPages = Math.ceil(filteredResults.length / itemsPerPage);

  const getStatusIcon = (status) => {
    switch (status) {
      case 'Match':
        return <CheckCircle className="icon success" />;
      case 'Mismatch':
        return <XCircle className="icon error" />;
      case 'Untraceable':
        return <AlertTriangle className="icon warning" />;
      default:
        return <AlertTriangle className="icon" />;
    }
  };

  const getStatusClass = (status) => {
    return `status-badge ${status.toLowerCase()}`;
  };

  return (
    <div className="audit-results">
      <div className="results-header">
        <h3>🔍 Detailed Results</h3>
        <div className="results-controls">
          <div className="search-box">
            <Search className="icon" />
            <input
              type="text"
              placeholder="Search results..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          <div className="filter-box">
            <Filter className="icon" />
            <select
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
            >
              <option value="all">All Status</option>
              <option value="Match">Matches</option>
              <option value="Mismatch">Mismatches</option>
              <option value="Untraceable">Untraceable</option>
            </select>
          </div>
        </div>
      </div>

      <div className="results-info">
        <p>Showing {paginatedResults.length} of {filteredResults.length} results</p>
      </div>

      <div className="results-table">
        {paginatedResults.map((result, index) => (
          <div key={index} className="result-card">
            <div className="result-header">
              <div className="slide-info">
                <span className="slide-number">Slide {result.slide}</span>
                <span className={getStatusClass(result.status)}>
                  {getStatusIcon(result.status)}
                  {result.status}
                </span>
              </div>
            </div>

            <div className="result-content">
              <div className="text-content">
                <h4>PPT Text:</h4>
                <p className="ppt-text">"{result.text}"</p>
                {result.context && (
                  <p className="context">Context: {result.context}</p>
                )}
              </div>

              <div className="values-comparison">
                <div className="value-item">
                  <label>PPT Value:</label>
                  <span className="value">{result.ppt_value?.toLocaleString() || 'N/A'}</span>
                </div>
                
                {result.excel_value !== null && (
                  <div className="value-item">
                    <label>Excel Value:</label>
                    <span className="value">{result.excel_value?.toLocaleString() || 'N/A'}</span>
                  </div>
                )}

                {result.suggested_fix && (
                  <div className="value-item suggested">
                    <label>Suggested Fix:</label>
                    <span className="value suggested-value">{result.suggested_fix}</span>
                  </div>
                )}
              </div>

              {result.excel_sheet && (
                <div className="excel-info">
                  <h4>Excel Source:</h4>
                  <p>
                    <strong>File:</strong> {result.excel_file}<br />
                    <strong>Sheet:</strong> {result.excel_sheet}<br />
                    <strong>Cell:</strong> {result.cell}
                  </p>
                </div>
              )}

              <div className="reasoning">
                <h4>Analysis:</h4>
                <p>{result.reasoning}</p>
                {result.confidence && (
                  <div className="confidence">
                    <span>Confidence: {(result.confidence * 100).toFixed(1)}%</span>
                  </div>
                )}
              </div>
            </div>
          </div>
        ))}
      </div>

      {totalPages > 1 && (
        <div className="pagination">
          <button
            onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
            disabled={currentPage === 1}
          >
            Previous
          </button>
          
          <span className="page-info">
            Page {currentPage} of {totalPages}
          </span>
          
          <button
            onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
            disabled={currentPage === totalPages}
          >
            Next
          </button>
        </div>
      )}
    </div>
  );
};

export default AuditResults;