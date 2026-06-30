import { NextRequest, NextResponse } from 'next/server';
import nodemailer from 'nodemailer';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    
    const smtpHost = formData.get('smtpHost') as string | null;
    const smtpPort = formData.get('smtpPort') as string | null;
    const smtpUser = formData.get('smtpUser') as string | null;
    const smtpPass = formData.get('smtpPass') as string | null;
    const emailSubject = formData.get('emailSubject') as string | null;
    const emailBody = formData.get('emailBody') as string | null;
    const recipientEmail = formData.get('recipientEmail') as string | null;
    
    const attachmentBlobs = formData.getAll('attachmentBlobs') as File[];
    const attachmentNames = formData.getAll('attachmentNames') as string[];
    const extraImage = formData.get('extraImage') as File | null;

    if (!smtpHost || !smtpPort || !smtpUser || !smtpPass || !emailSubject || !emailBody || !recipientEmail) {
      return NextResponse.json({ error: 'Missing required fields' }, { status: 400 });
    }

    // Configure Transporter
    const transporter = nodemailer.createTransport({
      host: smtpHost.trim(),
      port: parseInt(smtpPort, 10),
      secure: parseInt(smtpPort, 10) === 465,
      auth: {
        user: smtpUser.trim(),
        pass: smtpPass.replace(/\s+/g, ''),
      },
    });

    const mailAttachments = [];

    // Attach regular PDFs
    for (let i = 0; i < attachmentBlobs.length; i++) {
       const blob = attachmentBlobs[i];
       let name = attachmentNames[i];
       
       try { name = decodeURIComponent(name); } catch(e) {}
       
       const buffer = Buffer.from(await blob.arrayBuffer());
       mailAttachments.push({
         filename: name,
         content: buffer
       });
    }

    // Convert plain text body to HTML to preserve line breaks and allow embedding
    let finalBodyHTML = emailBody.replace(/\n/g, '<br/>');

    // Attach extra image inline
    if (extraImage) {
      let extraImageName = extraImage.name; 
      try { extraImageName = decodeURIComponent(extraImageName); } catch(e) {}
      
      const extraImageBuffer = Buffer.from(await extraImage.arrayBuffer());
      
      mailAttachments.push({
        filename: extraImageName,
        content: extraImageBuffer,
        cid: 'embedded-extra-image'
      });

      finalBodyHTML += `<br/><br/><img src="cid:embedded-extra-image" alt="${extraImageName}" style="max-width: 100%; height: auto; display: block; margin-top: 20px;" />`;
    }

    // Send Email
    await transporter.sendMail({
      from: smtpUser,
      to: recipientEmail,
      subject: emailSubject,
      html: finalBodyHTML,
      attachments: mailAttachments,
    });

    return NextResponse.json({ success: true });
  } catch (error: any) {
    console.error('Error sending single email:', error);
    return NextResponse.json({ error: error.message || 'Internal server error' }, { status: 500 });
  }
}
