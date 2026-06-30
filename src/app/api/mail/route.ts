import { NextRequest, NextResponse } from 'next/server';
import * as xlsx from 'xlsx';
import nodemailer from 'nodemailer';
import path from 'path';
import fs from 'fs';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    
    const dataFile = formData.get('dataFile') as File | null;
    const emailColumn = formData.get('emailColumn') as string | null;
    const attachmentColumn = formData.get('attachmentColumn') as string | null;
    const smtpHost = formData.get('smtpHost') as string | null;
    const smtpPort = formData.get('smtpPort') as string | null;
    const smtpUser = formData.get('smtpUser') as string | null;
    const smtpPass = formData.get('smtpPass') as string | null;
    const attachmentFolder = formData.get('attachmentFolder') as string | null;
    const emailSubject = formData.get('emailSubject') as string | null;
    const emailBody = formData.get('emailBody') as string | null;
    
    // Read the blobs and names separately to prevent multipart Thai corruption
    const attachmentBlobs = formData.getAll('attachmentBlobs') as File[];
    const attachmentNames = formData.getAll('attachmentNames') as string[];
    
    const attachmentFiles = attachmentBlobs.map((blob, index) => ({
      blob,
      name: attachmentNames[index] || `file${index}.pdf`
    }));

    const extraImage = formData.get('extraImage') as File | null;

    if (!dataFile || !emailColumn || !attachmentColumn || !smtpHost || !smtpPort || !smtpUser || !smtpPass || !emailSubject || !emailBody) {
      return NextResponse.json({ error: 'Missing required fields' }, { status: 400 });
    }

    // Configure Transporter
    const transporter = nodemailer.createTransport({
      host: smtpHost.trim(),
      port: parseInt(smtpPort, 10),
      secure: parseInt(smtpPort, 10) === 465, // true for 465, false for other ports
      auth: {
        user: smtpUser.trim(),
        pass: smtpPass.replace(/\s+/g, ''), // App passwords often get pasted with spaces
      },
    });

    // Verify connection config
    try {
      await transporter.verify();
    } catch (verifyError: any) {
      console.error('SMTP Connection Error:', verifyError);
      return NextResponse.json({ 
        error: `SMTP Error: ${verifyError.message || 'Failed to connect to SMTP server. Check your credentials and port.'}` 
      }, { status: 400 });
    }

    // Read Excel/CSV data
    const dataBuffer = Buffer.from(await dataFile.arrayBuffer());
    const workbook = xlsx.read(dataBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = xlsx.utils.sheet_to_json<any>(worksheet);

    // Buffer the extra image once to avoid reading it inside the loop
    let extraImageBuffer: Buffer | null = null;
    let extraImageName: string | null = null;
    if (extraImage) {
      extraImageBuffer = Buffer.from(await extraImage.arrayBuffer());
      try {
        extraImageName = decodeURIComponent(extraImage.name);
      } catch (e) {
        extraImageName = extraImage.name;
      }
    }

    let sentCount = 0;

    for (const row of jsonData) {
      const recipientEmail = row[emailColumn];
      if (!recipientEmail || typeof recipientEmail !== 'string' || !recipientEmail.includes('@')) {
        continue; // Skip invalid emails
      }

      // 1. Process Subject and Body replacements
      let finalSubject = emailSubject;
      let finalBody = emailBody;

      // Replace $ColumnName with row[ColumnName]
      Object.keys(row).forEach((key) => {
        // Use a global regex to replace all instances of $Key
        // Escape the key to prevent regex errors with special characters in column names
        const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(`\\$${escapedKey}`, 'g');
        const value = row[key] ? row[key].toString() : '';
        
        finalSubject = finalSubject.replace(regex, value);
        finalBody = finalBody.replace(regex, value);
      });

      // 2. Process Attachments (supports comma-separated list like x,y,z)
      let rawFilenamesStr = row[attachmentColumn];
      const mailAttachments = [];
      
      if (rawFilenamesStr) {
        const rawFilenames = rawFilenamesStr.toString().split(',');
        
        for (let rawName of rawFilenames) {
          rawName = rawName.trim();
          if (!rawName) continue;
          
          if (!rawName.toLowerCase().endsWith('.pdf')) {
            rawName += '.pdf';
          }
          
          // Sanitize same way as slicer
          const safeFilename = rawName.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_');
          const normalizedTarget = safeFilename.replace(/\s+/g, ''); // Strip spaces for robust matching
          
          // Find the matching file from the uploaded folder
          const matchingFile = attachmentFiles.find(f => {
            const normalizedSource = f.name.replace(/\s+/g, '');
            return normalizedSource === normalizedTarget;
          });

          if (matchingFile) {
            const fileBuffer = Buffer.from(await matchingFile.blob.arrayBuffer());
            mailAttachments.push({
              filename: safeFilename,
              content: fileBuffer
            });
          } else {
            console.warn(`Attachment not found in uploaded folder: ${safeFilename}`);
          }
        }
      }

      // Convert plain text body to HTML to preserve line breaks and allow embedding
      let finalBodyHTML = finalBody.replace(/\n/g, '<br/>');

      // Append the extra picture attachment as an INLINE embedded image if it exists
      if (extraImageBuffer && extraImageName) {
        mailAttachments.push({
          filename: extraImageName,
          content: extraImageBuffer,
          cid: 'embedded-extra-image' // Unique identifier for embedding
        });

        // Append the image tag to the end of the email body
        finalBodyHTML += `<br/><br/><img src="cid:embedded-extra-image" alt="${extraImageName}" style="max-width: 100%; height: auto; display: block; margin-top: 20px;" />`;
      }

      // 3. Send Email
      await transporter.sendMail({
        from: smtpUser,
        to: recipientEmail,
        subject: finalSubject,
        html: finalBodyHTML, // Send as HTML
        attachments: mailAttachments,
      });

      sentCount++;
    }

    return NextResponse.json({ success: true, sentCount });
  } catch (error: any) {
    console.error('Error in mailing system:', error);
    return NextResponse.json({ error: error.message || 'Internal server error' }, { status: 500 });
  }
}
