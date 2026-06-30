'use client';

import { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { PDFDocument } from 'pdf-lib';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';

export default function Home() {
  const [activeTab, setActiveTab] = useState<'slicer' | 'mailer'>('slicer');

  // --- Slicer State ---
  const [mainPdf, setMainPdf] = useState<File | null>(null);
  const [slicerDataFile, setSlicerDataFile] = useState<File | null>(null);
  const [slicerColumns, setSlicerColumns] = useState<string[]>([]);
  const [slicerData, setSlicerData] = useState<any[]>([]);
  const [slicerColumnName, setSlicerColumnName] = useState('');
  const [pagesPerSlice, setPagesPerSlice] = useState(1);
  const [slicingProgress, setSlicingProgress] = useState(0);
  const [isSlicing, setIsSlicing] = useState(false);
  const [slicerStatus, setSlicerStatus] = useState<{ type: 'success' | 'error'; message: string } | null>(null);

  // --- Mailer State ---
  const [mailerDataFile, setMailerDataFile] = useState<File | null>(null);
  const [mailerColumns, setMailerColumns] = useState<string[]>([]);
  const [mailerData, setMailerData] = useState<any[]>([]);
  const [selectedRows, setSelectedRows] = useState<Set<number>>(new Set());
  const [emailColumn, setEmailColumn] = useState('');
  const [attachmentColumn, setAttachmentColumn] = useState('');
  
  // SMTP Settings
  const [smtpHost, setSmtpHost] = useState('smtp.gmail.com');
  const [smtpPort, setSmtpPort] = useState(465);
  const [smtpUser, setSmtpUser] = useState('');
  const [smtpPass, setSmtpPass] = useState('');
  const [rememberMe, setRememberMe] = useState(false);
  
  // Load saved SMTP credentials on mount
  useEffect(() => {
    const savedUser = localStorage.getItem('smtpUser');
    const savedPass = localStorage.getItem('smtpPass');
    if (savedUser && savedPass) {
      setSmtpUser(savedUser);
      setSmtpPass(savedPass);
      setRememberMe(true);
    }
  }, []);

  // Save/Remove SMTP credentials when changed
  useEffect(() => {
    if (rememberMe) {
      localStorage.setItem('smtpUser', smtpUser);
      localStorage.setItem('smtpPass', smtpPass);
    } else {
      localStorage.removeItem('smtpUser');
      localStorage.removeItem('smtpPass');
    }
  }, [rememberMe, smtpUser, smtpPass]);
  
  // Mail Content
  const [attachmentFiles, setAttachmentFiles] = useState<FileList | null>(null);
  const [extraImage, setExtraImage] = useState<File | null>(null);
  const [emailSubject, setEmailSubject] = useState('');
  const [emailBody, setEmailBody] = useState('');

  const [isMailing, setIsMailing] = useState(false);
  const [mailerStatus, setMailerStatus] = useState<{ type: 'success' | 'error'; message: string } | null>(null);
  const [mailingLogs, setMailingLogs] = useState<{email: string, status: 'success' | 'error', message: string}[]>([]);
  const [progress, setProgress] = useState({ current: 0, total: 0 });

  const [isFetchingColumns, setIsFetchingColumns] = useState(false);

  // --- Autocomplete State ---
  const [autoComplete, setAutoComplete] = useState<{
    visible: boolean;
    filter: string;
    target: 'subject' | 'body' | null;
    startIndex: number;
    options: string[];
    selectedIndex: number;
  }>({
    visible: false,
    filter: '',
    target: null,
    startIndex: -1,
    options: [],
    selectedIndex: 0,
  });

  const subjectRef = useRef<HTMLInputElement>(null);
  const bodyRef = useRef<HTMLTextAreaElement>(null);

  const handleAutocompleteCheck = (text: string, cursorPosition: number, target: 'subject' | 'body') => {
    const textBeforeCursor = text.slice(0, cursorPosition);
    const match = textBeforeCursor.match(/\$([a-zA-Z0-9_ก-๙-]*)$/);

    if (match) {
      const filter = match[1];
      const startIndex = cursorPosition - filter.length - 1;
      
      const filteredOptions = mailerColumns.filter(col => 
        col.toLowerCase().includes(filter.toLowerCase())
      );

      if (filteredOptions.length > 0) {
        setAutoComplete({
          visible: true,
          filter,
          target,
          startIndex,
          options: filteredOptions,
          selectedIndex: 0,
        });
        return;
      }
    }
    
    setAutoComplete(prev => ({ ...prev, visible: false }));
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>, target: 'subject' | 'body') => {
    const val = e.target.value;
    if (target === 'subject') setEmailSubject(val);
    else setEmailBody(val);
    
    handleAutocompleteCheck(val, e.target.selectionStart || 0, target);
  };

  const insertSuggestion = (suggestion: string) => {
    if (!autoComplete.target) return;
    
    const text = autoComplete.target === 'subject' ? emailSubject : emailBody;
    const ref = autoComplete.target === 'subject' ? subjectRef.current : bodyRef.current;
    
    const before = text.slice(0, autoComplete.startIndex);
    const after = text.slice(autoComplete.startIndex + autoComplete.filter.length + 1);
    
    const newText = before + '$' + suggestion + after;
    
    if (autoComplete.target === 'subject') setEmailSubject(newText);
    else setEmailBody(newText);
    
    setAutoComplete(prev => ({ ...prev, visible: false }));
    
    setTimeout(() => {
      if (ref) {
        ref.focus();
        const newCursorPos = before.length + 1 + suggestion.length;
        ref.setSelectionRange(newCursorPos, newCursorPos);
      }
    }, 0);
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>) => {
    if (autoComplete.visible) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setAutoComplete(prev => ({ ...prev, selectedIndex: Math.min(prev.selectedIndex + 1, prev.options.length - 1) }));
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        setAutoComplete(prev => ({ ...prev, selectedIndex: Math.max(prev.selectedIndex - 1, 0) }));
      } else if (e.key === 'Enter' || e.key === 'Tab') {
        e.preventDefault();
        insertSuggestion(autoComplete.options[autoComplete.selectedIndex]);
      } else if (e.key === 'Escape') {
        setAutoComplete(prev => ({ ...prev, visible: false }));
      }
    }
  };

  const handleDataFileChange = async (e: React.ChangeEvent<HTMLInputElement>, type: 'slicer' | 'mailer') => {
    const file = e.target.files?.[0] || null;
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);
      const cols = Object.keys(data[0] || {});

      if (type === 'slicer') {
        setSlicerDataFile(file);
        setSlicerData(data);
        setSlicerColumns(cols);
        if (cols.length > 0) setSlicerColumnName(cols[0]);
      } else {
        setMailerDataFile(file);
        setMailerData(data);
        setMailerColumns(cols);
        setSelectedRows(new Set(data.map((_, i) => i)));
        if (cols.length > 0) {
          setEmailColumn(cols[0]);
          setAttachmentColumn(cols[0]);
        }
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleSliceSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!mainPdf || !slicerDataFile || !slicerColumnName) {
      setSlicerStatus({ type: 'error', message: 'Please provide all files and select a naming column' });
      return;
    }

    setIsSlicing(true);
    setSlicerStatus(null);
    setSlicingProgress(0);

    try {
      const pdfBytes = await mainPdf.arrayBuffer();
      const pdfDoc = await PDFDocument.load(pdfBytes);
      const numPages = pdfDoc.getPageCount();
      const numSlices = Math.ceil(numPages / pagesPerSlice);

      if (numSlices !== slicerData.length) {
        throw new Error(`Page count mismatch! PDF has ${numPages} pages (which makes ${numSlices} slices of ${pagesPerSlice} pages), but Excel has ${slicerData.length} rows.`);
      }

      let dirHandle: any = null;
      if ('showDirectoryPicker' in window) {
        try {
          dirHandle = await (window as any).showDirectoryPicker({
            mode: 'readwrite',
          });
        } catch (err) {
          // User cancelled the directory picker
          setIsSlicing(false);
          setSlicingProgress(0);
          return;
        }
      }

      const zip = new JSZip();

      for (let i = 0; i < numSlices; i++) {
        const row = slicerData[i];
        let filename = String(row[slicerColumnName] || `Slice_${i + 1}`).replace(/[\/\\?%*:|"<>]/g, '_');
        if (!filename.toLowerCase().endsWith('.pdf')) {
          filename += '.pdf';
        }

        const newPdf = await PDFDocument.create();
        
        const pagesToCopy = [];
        for (let p = 0; p < pagesPerSlice; p++) {
          const pageIndex = i * pagesPerSlice + p;
          if (pageIndex < numPages) {
            pagesToCopy.push(pageIndex);
          }
        }

        const copiedPages = await newPdf.copyPages(pdfDoc, pagesToCopy);
        copiedPages.forEach((page) => newPdf.addPage(page));

        const slicedPdfBytes = await newPdf.save();
        
        if (dirHandle) {
          const fileHandle = await dirHandle.getFileHandle(filename, { create: true });
          const writable = await fileHandle.createWritable();
          await writable.write(slicedPdfBytes);
          await writable.close();
        } else {
          zip.file(filename, slicedPdfBytes);
        }

        setSlicingProgress(Math.round(((i + 1) / numSlices) * 100));
      }

      if (!dirHandle) {
        setSlicerStatus({ type: 'success', message: 'Generating ZIP file...' });
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        saveAs(zipBlob, 'Sliced_PDFs.zip');
        setSlicerStatus({ type: 'success', message: 'Success! Sliced PDFs downloaded as a ZIP file (Folder saving not supported in your browser).' });
      } else {
        setSlicerStatus({ type: 'success', message: 'Success! Sliced PDFs saved directly to your selected folder.' });
      }
    } catch (error: any) {
      setSlicerStatus({ type: 'error', message: error.message || 'An error occurred during slicing' });
    } finally {
      setIsSlicing(false);
      setSlicingProgress(0);
    }
  };

  const handleMailSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!mailerDataFile || !emailColumn || !attachmentColumn || !smtpHost || !smtpUser || !smtpPass) {
      setMailerStatus({ type: 'error', message: 'Please fill in all required SMTP and Data fields.' });
      return;
    }

    const selectedArray = Array.from(selectedRows);
    if (selectedArray.length === 0) {
      setMailerStatus({ type: 'error', message: 'Please select at least one recipient.' });
      return;
    }

    setIsMailing(true);
    setMailerStatus(null);
    setMailingLogs([]);
    setProgress({ current: 0, total: selectedArray.length });

    let sentCount = 0;

    for (const index of selectedArray) {
      const row = mailerData[index];
      const recipientEmail = row[emailColumn];
      
      if (!recipientEmail || typeof recipientEmail !== 'string' || !recipientEmail.includes('@')) {
        setMailingLogs(prev => [{ email: 'Unknown', status: 'error', message: 'Invalid email address' }, ...prev]);
        setProgress(prev => ({ ...prev, current: prev.current + 1 }));
        continue;
      }

      let finalSubject = emailSubject;
      let finalBody = emailBody;
      Object.keys(row).forEach((key) => {
        const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(`\\$${escapedKey}`, 'g');
        const value = row[key] ? row[key].toString() : '';
        finalSubject = finalSubject.replace(regex, value);
        finalBody = finalBody.replace(regex, value);
      });

      const formData = new FormData();
      formData.append('smtpHost', smtpHost);
      formData.append('smtpPort', smtpPort.toString());
      formData.append('smtpUser', smtpUser);
      formData.append('smtpPass', smtpPass);
      formData.append('emailSubject', finalSubject);
      formData.append('emailBody', finalBody);
      formData.append('recipientEmail', recipientEmail);

      let rawFilenamesStr = row[attachmentColumn];
      if (rawFilenamesStr && attachmentFiles) {
        const rawFilenames = rawFilenamesStr.toString().split(',');
        for (let rawName of rawFilenames) {
          rawName = rawName.trim();
          if (!rawName) continue;
          if (!rawName.toLowerCase().endsWith('.pdf')) rawName += '.pdf';
          
          const safeFilename = rawName.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_');
          const normalizedTarget = safeFilename.replace(/\s+/g, '');
          
          const matchingFile = Array.from(attachmentFiles).find(f => {
            const normalizedSource = f.name.replace(/\s+/g, '');
            return normalizedSource === normalizedTarget;
          });

          if (matchingFile) {
            formData.append('attachmentBlobs', matchingFile, `file`);
            formData.append('attachmentNames', encodeURIComponent(matchingFile.name));
          }
        }
      }

      if (extraImage) {
        formData.append('extraImage', extraImage, encodeURIComponent(extraImage.name));
      }

      try {
        const res = await fetch('/api/mail/send-single', {
          method: 'POST',
          body: formData,
        });

        const data = await res.json();
        if (res.ok) {
          setMailingLogs(prev => [{ email: recipientEmail, status: 'success', message: 'Sent successfully' }, ...prev]);
          sentCount++;
        } else {
          setMailingLogs(prev => [{ email: recipientEmail, status: 'error', message: data.error || 'Failed' }, ...prev]);
        }
      } catch (err: any) {
        setMailingLogs(prev => [{ email: recipientEmail, status: 'error', message: 'Network error' }, ...prev]);
      }

      setProgress(prev => ({ ...prev, current: prev.current + 1 }));
    }

    setMailerStatus({ type: 'success', message: `Finished. Sent ${sentCount} out of ${selectedArray.length} emails.` });
    setIsMailing(false);
  };

  return (
    <div className="dashboard-container">
      <div className="dashboard-header">
        <h1>{activeTab === 'slicer' ? 'PDF Slicer' : 'Mailing System'}</h1>
        <p>
          {activeTab === 'slicer' 
            ? 'Slice your main PDF into multiple files based on Excel/CSV data.' 
            : 'Send personalized emails with attachments based on Excel/CSV data.'}
        </p>
      </div>

      <div className="tabs">
        <button 
          className={`tab-btn ${activeTab === 'slicer' ? 'active' : ''}`}
          onClick={() => setActiveTab('slicer')}
        >
          PDF Slicer
        </button>
        <button 
          className={`tab-btn ${activeTab === 'mailer' ? 'active' : ''}`}
          onClick={() => setActiveTab('mailer')}
        >
          Mailing System
        </button>
      </div>

      {activeTab === 'slicer' && (
        <form onSubmit={handleSliceSubmit}>
          <div className="form-grid">
            <div className="form-group full-width">
              <label htmlFor="mainPdf">Main PDF File</label>
              <input
                type="file"
                id="mainPdf"
                accept="application/pdf"
                onChange={(e) => setMainPdf(e.target.files?.[0] || null)}
                required
              />
            </div>

            <div className="form-group full-width">
              <label htmlFor="slicerDataFile">Data File (Excel/CSV)</label>
              <input
                type="file"
                id="slicerDataFile"
                accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
                onChange={(e) => handleDataFileChange(e, 'slicer')}
                required
              />
            </div>

            <div className="form-group full-width">
              <label htmlFor="slicerColumnName">Filename Column</label>
              <select
                id="slicerColumnName"
                value={slicerColumnName}
                onChange={(e) => setSlicerColumnName(e.target.value)}
                required
                disabled={slicerColumns.length === 0}
              >
                {slicerColumns.length === 0 && <option value="">-- Select Data File First --</option>}
                {slicerColumns.map((col, idx) => (
                  <option key={idx} value={col}>{col}</option>
                ))}
              </select>
            </div>

            <div className="form-group full-width">
              <label htmlFor="pagesPerSlice">Pages per Slice</label>
              <input
                type="number"
                id="pagesPerSlice"
                min="1"
                value={pagesPerSlice}
                onChange={(e) => setPagesPerSlice(parseInt(e.target.value) || 1)}
                required
              />
            </div>
          </div>

          <button type="submit" className="submit-btn" disabled={isSlicing || !slicerColumnName}>
            {isSlicing ? (
              <>
                <span className="loader"></span> 
                {slicingProgress > 0 ? `Slicing... ${slicingProgress}%` : 'Processing...'}
              </>
            ) : 'Slice PDF'}
          </button>

          {slicerStatus && (
            <div className={`status-message ${slicerStatus.type}`}>{slicerStatus.message}</div>
          )}
        </form>
      )}

      {/* --- MAILER TAB --- */}
      {activeTab === 'mailer' && (
        <form onSubmit={handleMailSubmit}>
          <div className="form-grid">
            <div className="form-group full-width">
              <label htmlFor="mailerDataFile">Data File (Excel/CSV)</label>
              <input
                type="file"
                id="mailerDataFile"
                accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
                onChange={(e) => handleDataFileChange(e, 'mailer')}
                required
              />
            </div>

            <div className="form-group">
              <label htmlFor="emailColumn">
                Email Address Column {isFetchingColumns && <span className="loader small"></span>}
              </label>
              <select
                id="emailColumn"
                value={emailColumn}
                onChange={(e) => setEmailColumn(e.target.value)}
                required
                disabled={mailerColumns.length === 0 || isFetchingColumns}
              >
                {mailerColumns.length === 0 && <option value="">-- Select Data File First --</option>}
                {mailerColumns.map((col, idx) => (
                  <option key={idx} value={col}>{col}</option>
                ))}
              </select>
            </div>

            <div className="form-group">
              <label htmlFor="attachmentColumn">
                Attachment Name Column {isFetchingColumns && <span className="loader small"></span>}
              </label>
              <select
                id="attachmentColumn"
                value={attachmentColumn}
                onChange={(e) => setAttachmentColumn(e.target.value)}
                required
                disabled={mailerColumns.length === 0 || isFetchingColumns}
              >
                {mailerColumns.length === 0 && <option value="">-- Select Data File First --</option>}
                {mailerColumns.map((col, idx) => (
                  <option key={idx} value={col}>{col}</option>
                ))}
              </select>
            </div>
            
            {mailerData.length > 0 && (
              <div className="form-group full-width" style={{ marginTop: '0.5rem' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.5rem' }}>
                  <h3 style={{ color: 'var(--text-secondary)', fontSize: '1rem', margin: 0 }}>Select Recipients ({selectedRows.size} of {mailerData.length})</h3>
                  <button 
                    type="button" 
                    className="tab-btn active" 
                    style={{ padding: '0.2rem 0.5rem', fontSize: '0.8rem', minHeight: 'auto' }}
                    onClick={() => {
                      if (selectedRows.size === mailerData.length) {
                        setSelectedRows(new Set());
                      } else {
                        setSelectedRows(new Set(mailerData.map((_, i) => i)));
                      }
                    }}
                  >
                    {selectedRows.size === mailerData.length ? 'Deselect All' : 'Select All'}
                  </button>
                </div>
                <div style={{ maxHeight: '250px', overflow: 'auto', border: '1px solid var(--card-border)', borderRadius: '8px', background: 'var(--bg-secondary)' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85rem', textAlign: 'left', whiteSpace: 'nowrap' }}>
                    <thead style={{ position: 'sticky', top: 0, background: 'var(--card-bg)', boxShadow: '0 1px 2px rgba(0,0,0,0.1)', zIndex: 10 }}>
                      <tr>
                        <th style={{ padding: '0.5rem', width: '40px', textAlign: 'center', position: 'sticky', left: 0, background: '#fee2e2', zIndex: 11, borderRight: '1px solid var(--card-border)' }}>#</th>
                        {mailerColumns.map((col, idx) => (
                          <th key={idx} style={{ padding: '0.5rem' }}>{col}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {mailerData.map((row, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--card-border)' }}>
                          <td style={{ padding: '0.5rem', textAlign: 'center', position: 'sticky', left: 0, background: '#fef2f2', borderRight: '1px solid var(--card-border)', zIndex: 1 }}>
                            <input 
                              type="checkbox" 
                              checked={selectedRows.has(idx)}
                              onChange={(e) => {
                                const newSet = new Set(selectedRows);
                                if (e.target.checked) newSet.add(idx);
                                else newSet.delete(idx);
                                setSelectedRows(newSet);
                              }}
                            />
                          </td>
                          {mailerColumns.map((col, cIdx) => (
                            <td key={cIdx} style={{ padding: '0.5rem' }}>{row[col] || '-'}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div className="form-group full-width">
              <label htmlFor="attachmentFolder">Browse Attachment Folder</label>
              <input
                type="file"
                id="attachmentFolder"
                /* @ts-expect-error webkitdirectory is non-standard but widely supported */
                webkitdirectory=""
                directory=""
                onChange={(e) => setAttachmentFiles(e.target.files)}
                required
              />
            </div>

            <div className="form-group full-width">
              <label htmlFor="extraImage">Additional Picture Attachment (Optional)</label>
              <input
                type="file"
                id="extraImage"
                accept="image/*"
                onChange={(e) => setExtraImage(e.target.files?.[0] || null)}
              />
            </div>
            
            <div className="form-group full-width" style={{ marginTop: '1rem', borderTop: '1px solid var(--card-border)', paddingTop: '1rem' }}>
              <h3 style={{ color: 'var(--text-secondary)', marginBottom: '0.5rem', fontSize: '1.1rem' }}>Email Content</h3>
              <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '1rem' }}>
                Tip: Type <code style={{ color: 'var(--accent-color)' }}>$</code> to auto-complete your column variables. Works for both Thai and English.
              </p>
            </div>

            <div className="form-group full-width">
              <label htmlFor="emailSubject">Subject</label>
              <div className="input-wrapper">
                <input 
                  type="text" 
                  id="emailSubject" 
                  ref={subjectRef}
                  placeholder="Welcome $Name from $Institute" 
                  value={emailSubject} 
                  onChange={(e) => handleInputChange(e, 'subject')} 
                  onKeyDown={handleKeyDown}
                  onKeyUp={(e) => handleAutocompleteCheck((e.target as HTMLInputElement).value, (e.target as HTMLInputElement).selectionStart || 0, 'subject')}
                  onClick={(e) => handleAutocompleteCheck((e.target as HTMLInputElement).value, (e.target as HTMLInputElement).selectionStart || 0, 'subject')}
                  required 
                />
                {autoComplete.visible && autoComplete.target === 'subject' && (
                  <div className="autocomplete-dropdown">
                    {autoComplete.options.map((opt, idx) => (
                      <div 
                        key={idx} 
                        className={`autocomplete-item ${idx === autoComplete.selectedIndex ? 'selected' : ''}`}
                        onClick={() => insertSuggestion(opt)}
                      >
                        ${opt}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
            
            <div className="form-group full-width">
              <label htmlFor="emailBody">Message Body</label>
              <div className="input-wrapper">
                <textarea 
                  id="emailBody" 
                  ref={bodyRef}
                  placeholder="Hello $Name,&#10;&#10;Here is your file: $File" 
                  value={emailBody} 
                  onChange={(e) => handleInputChange(e, 'body')} 
                  onKeyDown={handleKeyDown}
                  onKeyUp={(e) => handleAutocompleteCheck((e.target as HTMLTextAreaElement).value, (e.target as HTMLTextAreaElement).selectionStart || 0, 'body')}
                  onClick={(e) => handleAutocompleteCheck((e.target as HTMLTextAreaElement).value, (e.target as HTMLTextAreaElement).selectionStart || 0, 'body')}
                  required
                ></textarea>
                {autoComplete.visible && autoComplete.target === 'body' && (
                  <div className="autocomplete-dropdown">
                    {autoComplete.options.map((opt, idx) => (
                      <div 
                        key={idx} 
                        className={`autocomplete-item ${idx === autoComplete.selectedIndex ? 'selected' : ''}`}
                        onClick={() => insertSuggestion(opt)}
                      >
                        ${opt}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
            
            <div className="form-group full-width" style={{ marginTop: '1rem', borderTop: '1px solid var(--card-border)', paddingTop: '1rem' }}>
              <h3 style={{ color: 'var(--text-secondary)', marginBottom: '0.5rem', fontSize: '1.1rem' }}>SMTP Configuration</h3>
            </div>

            <div className="form-group">
              <label htmlFor="smtpHost">SMTP Host</label>
              <input type="text" id="smtpHost" placeholder="e.g., smtp.gmail.com" value={smtpHost} onChange={(e) => setSmtpHost(e.target.value)} required />
            </div>
            <div className="form-group">
              <label htmlFor="smtpPort">SMTP Port</label>
              <input type="number" id="smtpPort" placeholder="465" value={smtpPort} onChange={(e) => setSmtpPort(parseInt(e.target.value) || 465)} required />
            </div>
            <div className="form-group">
              <label htmlFor="smtpUser">Email / Username</label>
              <input type="text" id="smtpUser" placeholder="you@gmail.com" value={smtpUser} onChange={(e) => setSmtpUser(e.target.value)} required />
            </div>
            <div className="form-group">
              <label htmlFor="smtpPass">Password / App Password</label>
              <input type="password" id="smtpPass" value={smtpPass} onChange={(e) => setSmtpPass(e.target.value)} required />
            </div>
            
            <div className="form-group full-width" style={{ marginTop: '-0.5rem', marginBottom: '1rem' }}>
              <label style={{ display: 'flex', alignItems: 'center', cursor: 'pointer', gap: '8px', fontSize: '0.9rem', fontWeight: 'normal', color: 'var(--text-secondary)' }}>
                <input 
                  type="checkbox" 
                  checked={rememberMe} 
                  onChange={(e) => setRememberMe(e.target.checked)} 
                  style={{ width: 'auto', marginBottom: 0 }}
                />
                Remember Email & App Password
              </label>
            </div>

          </div>
          
          {isMailing && (
             <div className="form-group full-width" style={{ marginTop: '1rem', padding: '1rem', background: 'var(--bg-secondary)', borderRadius: '8px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '0.5rem' }}>
                   <strong>Sending Emails...</strong>
                   <span>{progress.current} / {progress.total}</span>
                </div>
                <div style={{ width: '100%', height: '8px', background: 'var(--card-border)', borderRadius: '4px', overflow: 'hidden' }}>
                   <div style={{ height: '100%', background: 'var(--accent-color)', width: `${progress.total > 0 ? (progress.current / progress.total) * 100 : 0}%`, transition: 'width 0.2s' }}></div>
                </div>
             </div>
          )}

          {mailingLogs.length > 0 && (
            <div className="form-group full-width" style={{ marginTop: '1rem' }}>
               <h3 style={{ fontSize: '1rem', marginBottom: '0.5rem' }}>Dispatch Logs</h3>
               <div style={{ maxHeight: '200px', overflowY: 'auto', border: '1px solid var(--card-border)', borderRadius: '8px', padding: '0.5rem', background: 'var(--bg-color)', fontSize: '0.8rem', fontFamily: 'monospace' }}>
                  {mailingLogs.map((log, idx) => (
                    <div key={idx} style={{ padding: '0.2rem 0', color: log.status === 'success' ? '#10b981' : '#ef4444' }}>
                      [{log.status === 'success' ? 'SUCCESS' : 'FAILED'}] {log.email}: {log.message}
                    </div>
                  ))}
               </div>
            </div>
          )}

          <button type="submit" className="submit-btn" disabled={isMailing || !emailColumn}>
            {isMailing ? <><span className="loader"></span> Sending Emails...</> : 'Send Emails'}
          </button>

          {mailerStatus && (
            <div className={`status-message ${mailerStatus.type}`}>{mailerStatus.message}</div>
          )}
        </form>
      )}
    </div>
  );
}
