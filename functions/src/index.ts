import * as functions from 'firebase-functions';
import * as admin from 'firebase-admin';
import * as nodemailer from 'nodemailer';

// Initialize Firebase Admin
if (!admin.apps.length) {
  admin.initializeApp();
}

// Target database specified in project configuration
const databaseId = process.env.FIRESTORE_DATABASE_ID || 'ai-studio-348fddfd-afda-4dd5-bc81-e0d8528324bd';
const db = admin.firestore();
// If custom databaseId is used in Firebase Admin:
// const db = admin.firestore(databaseId);

// Email Transporter setup
function getEmailTransporter() {
  const host = process.env.SMTP_HOST || functions.config().smtp?.host;
  const port = Number(process.env.SMTP_PORT || functions.config().smtp?.port || 587);
  const user = process.env.SMTP_USER || functions.config().smtp?.user;
  const pass = process.env.SMTP_PASS || functions.config().smtp?.pass;

  if (host && user && pass) {
    return nodemailer.createTransport({
      host,
      port,
      secure: port === 465,
      auth: { user, pass },
    });
  }

  // Fallback or test mode
  return null;
}

// Generate HTML email template for 3-day PM reminder
function generatePMEmailHtml(data: {
  orderId: string;
  title: string;
  equipmentName: string;
  equipmentId: string;
  factory?: string;
  dueDateStr: string;
  daysRemaining: number;
  priority: string;
  assignedTo?: string;
  description?: string;
  pmFrequency?: string;
}): string {
  const priorityColor =
    data.priority === 'urgent'
      ? '#e11d48'
      : data.priority === 'high'
      ? '#ea580c'
      : '#2563eb';

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #f8fafc; color: #1e293b; margin: 0; padding: 24px; }
    .card { background: #ffffff; border-radius: 12px; border: 1px solid #e2e8f0; max-width: 600px; margin: 0 auto; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .header { background: #0f172a; color: #ffffff; padding: 24px; text-align: left; }
    .badge { display: inline-block; padding: 4px 10px; font-size: 12px; font-weight: 700; border-radius: 9999px; text-transform: uppercase; letter-spacing: 0.5px; }
    .badge-alert { background: #fef3c7; color: #b45309; border: 1px solid #fde68a; }
    .content { padding: 24px; }
    .info-table { width: 100%; border-collapse: collapse; margin-top: 16px; margin-bottom: 20px; }
    .info-table td { padding: 10px 12px; border-bottom: 1px solid #f1f5f9; font-size: 14px; }
    .info-table td.label { font-weight: 600; color: #64748b; width: 35%; }
    .info-table td.value { color: #0f172a; font-weight: 500; }
    .button { display: inline-block; background-color: #2563eb; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 8px; font-weight: 600; font-size: 14px; margin-top: 12px; }
    .footer { padding: 16px 24px; background: #f8fafc; border-top: 1px solid #e2e8f0; font-size: 12px; color: #64748b; text-align: center; }
  </style>
</head>
<body>
  <div class="card">
    <div class="header">
      <div style="font-size: 12px; text-transform: uppercase; letter-spacing: 1px; color: #94a3b8; margin-bottom: 6px;">
        Hệ thống Quản lý Bảo trì CMMS
      </div>
      <h1 style="margin: 0; font-size: 20px; font-weight: 700; color: #ffffff;">
        Cảnh báo: Lịch bảo trì định kỳ (PM) sắp đến hạn
      </h1>
    </div>

    <div class="content">
      <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px;">
        <span class="badge badge-alert">
          ⏰ Còn đúng 3 ngày (Hạn: ${data.dueDateStr})
        </span>
        <span style="font-size: 12px; font-weight: 600; color: ${priorityColor}; text-transform: uppercase;">
          Ưu tiên: ${data.priority}
        </span>
      </div>

      <p style="font-size: 15px; line-height: 1.5; color: #334155; margin-top: 0;">
        Kính gửi Đội ngũ Kỹ thuật & Quản lý vận hành,<br/>
        Hệ thống phát hiện phiếu công việc bảo trì định kỳ (Preventive Maintenance) sau đây sẽ đến hạn trong <strong>3 ngày tới</strong>. Vui lòng chuẩn bị vật tư, nhân lực và phương án an toàn.
      </p>

      <table class="info-table">
        <tr>
          <td class="label">Phiếu công việc:</td>
          <td class="value"><strong>${data.title}</strong></td>
        </tr>
        <tr>
          <td class="label">Mã phiếu (ID):</td>
          <td class="value" style="font-family: monospace;">${data.orderId}</td>
        </tr>
        <tr>
          <td class="label">Thiết bị:</td>
          <td class="value">${data.equipmentName} (${data.equipmentId})</td>
        </tr>
        ${data.factory ? `
        <tr>
          <td class="label">Nhà máy / Vị trí:</td>
          <td class="value">${data.factory}</td>
        </tr>` : ''}
        <tr>
          <td class="label">Ngày đến hạn:</td>
          <td class="value" style="color: #b45309; font-weight: 700;">${data.dueDateStr}</td>
        </tr>
        ${data.pmFrequency ? `
        <tr>
          <td class="label">Chu kỳ bảo trì:</td>
          <td class="value">${data.pmFrequency}</td>
        </tr>` : ''}
        ${data.assignedTo ? `
        <tr>
          <td class="label">Phân công kỹ thuật:</td>
          <td class="value">${data.assignedTo}</td>
        </tr>` : ''}
        ${data.description ? `
        <tr>
          <td class="label">Mô tả công việc:</td>
          <td class="value">${data.description}</td>
        </tr>` : ''}
      </table>

      <div style="text-align: center; margin: 24px 0 12px 0;">
        <p style="font-size: 13px; color: #64748b; margin-bottom: 12px;">
          Vui lòng kiểm tra phương án thao tác, phiếu xin mở đường điện (nếu có) và chuẩn bị vật tư thay thế.
        </p>
      </div>
    </div>

    <div class="footer">
      Thông báo tự động từ Firebase Cloud Functions • Renewable CMMS & Asset Analytics<br/>
      Email được gửi đến kỹ thuật viên phụ trách và ban quản trị kỹ thuật.
    </div>
  </div>
</body>
</html>
`;
}

// Core business logic: Find PM tasks due in 3 days and send alerts
export async function processUpcomingPMAlerts() {
  const now = new Date();
  
  // Calculate target date window: exactly 3 days from today
  const targetDate = new Date();
  targetDate.setDate(now.getDate() + 3);
  const targetDateStr = targetDate.toISOString().split('T')[0]; // YYYY-MM-DD

  console.log(`Checking for PM work orders due on or before 3 days: target = ${targetDateStr}`);

  // Query preventive work orders that are not completed/cancelled
  const workOrdersSnapshot = await db
    .collection('workOrders')
    .where('type', '==', 'preventive')
    .get();

  const results: any[] = [];
  const transporter = getEmailTransporter();
  const defaultAlertEmail = process.env.DEFAULT_ALERT_EMAIL || 'sgm1707@gmail.com';

  for (const doc of workOrdersSnapshot.docs) {
    const data = doc.data();
    
    // Skip finished or cancelled work orders
    if (data.status === 'completed' || data.status === 'cancelled') {
      continue;
    }

    if (!data.dueDate) {
      continue;
    }

    const dueDateObj = new Date(data.dueDate);
    if (isNaN(dueDateObj.getTime())) {
      continue;
    }

    const dueDateStr = dueDateObj.toISOString().split('T')[0];
    
    // Calculate difference in whole days between today and dueDate
    const diffTime = dueDateObj.getTime() - now.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    // Alert condition: Exactly 3 days remaining (or within 3 days: 0 <= diffDays <= 3) and not yet notified
    const isDueIn3Days = dueDateStr === targetDateStr || (diffDays >= 1 && diffDays <= 3);

    if (isDueIn3Days) {
      // Check if already notified for the 3-day milestone
      if (data.notified3DaysBefore === true) {
        console.log(`Order ${doc.id} already notified. Skipping.`);
        continue;
      }

      // Resolve equipment details
      let eqName = 'Thiết bị';
      const eqId = Array.isArray(data.equipmentId) ? data.equipmentId[0] : data.equipmentId;
      if (eqId) {
        try {
          const eqDoc = await db.collection('equipment').doc(eqId).get();
          if (eqDoc.exists) {
            eqName = eqDoc.data()?.name || eqName;
          }
        } catch (e) {
          console.warn(`Could not load equipment ${eqId}`, e);
        }
      }

      // Determine recipient emails
      const recipients = new Set<string>();
      recipients.add(defaultAlertEmail);

      // If assignedTo is an email or username
      if (data.assignedTo) {
        if (data.assignedTo.includes('@')) {
          recipients.add(data.assignedTo);
        } else {
          // Look up user by displayName or uid
          try {
            const userSnap = await db.collection('users').where('displayName', '==', data.assignedTo).get();
            userSnap.forEach(uDoc => {
              const uEmail = uDoc.data()?.email;
              if (uEmail) recipients.add(uEmail);
            });
          } catch (e) {
            // Ignore lookup error
          }
        }
      }

      // If customer has an email
      if (data.customerId) {
        try {
          const custDoc = await db.collection('customers').doc(data.customerId).get();
          if (custDoc.exists && custDoc.data()?.email) {
            recipients.add(custDoc.data()!.email);
          }
        } catch (e) {
          // Ignore
        }
      }

      const recipientList = Array.from(recipients);
      const emailHtml = generatePMEmailHtml({
        orderId: doc.id,
        title: data.title || 'Bảo trì định kỳ',
        equipmentName: eqName,
        equipmentId: eqId || 'N/A',
        factory: data.factory,
        dueDateStr: dueDateObj.toLocaleDateString('vi-VN'),
        daysRemaining: diffDays,
        priority: data.priority || 'medium',
        assignedTo: data.assignedTo,
        description: data.description,
        pmFrequency: data.pmFrequency,
      });

      let emailStatus = 'simulated';
      let errorMsg = null;

      if (transporter) {
        try {
          const from = process.env.NOTIFICATION_EMAIL_FROM || '"Renewable CMMS" <notifications@cmms-system.com>';
          await transporter.sendMail({
            from,
            to: recipientList.join(', '),
            subject: `[CẢNH BÁO PM - CÒN ${diffDays} NGÀY] ${data.title} - ${eqName} (Hạn: ${dueDateObj.toLocaleDateString('vi-VN')})`,
            html: emailHtml,
          });
          emailStatus = 'sent';
          console.log(`Sent PM 3-day notification to ${recipientList.join(', ')} for work order ${doc.id}`);
        } catch (err: any) {
          console.error(`Failed to send email for work order ${doc.id}:`, err);
          emailStatus = 'error';
          errorMsg = err.message;
        }
      } else {
        console.log(`[SIMULATED] SMTP not configured. Would send email to: ${recipientList.join(', ')} for work order ${doc.id}`);
      }

      // Update work order document to prevent duplicate alerts
      await db.collection('workOrders').doc(doc.id).update({
        notified3DaysBefore: true,
        lastNotificationSentAt: admin.firestore.FieldValue.serverTimestamp(),
        lastNotificationStatus: emailStatus,
      });

      // Record in notifications collection for audit log
      const logRef = await db.collection('notifications').add({
        type: 'pm_3_days_alert',
        workOrderId: doc.id,
        workOrderTitle: data.title,
        equipmentId: eqId,
        equipmentName: eqName,
        dueDate: data.dueDate,
        daysRemaining: diffDays,
        recipients: recipientList,
        status: emailStatus,
        error: errorMsg,
        createdAt: admin.firestore.FieldValue.serverTimestamp(),
      });

      results.push({
        workOrderId: doc.id,
        title: data.title,
        equipmentName: eqName,
        dueDate: data.dueDate,
        recipients: recipientList,
        status: emailStatus,
        notificationLogId: logRef.id,
      });
    }
  }

  return {
    checkedAt: new Date().toISOString(),
    targetDateStr,
    alertsDispatched: results.length,
    details: results,
  };
}

// 1. Firebase Cloud Functions v2 / v1: Scheduled job running every day at 08:00 AM (Asia/Ho_Chi_Minh)
export const scheduledPMNotifications = functions.pubsub
  .schedule('0 8 * * *')
  .timeZone('Asia/Ho_Chi_Minh')
  .onRun(async (context) => {
    console.log('Starting daily PM 3-day notification scan...');
    try {
      const summary = await processUpcomingPMAlerts();
      console.log('Completed PM notification scan:', summary);
      return null;
    } catch (error) {
      console.error('Error running scheduled PM notifications:', error);
      throw error;
    }
  });

// 2. HTTPS Endpoint to trigger or test PM notifications on-demand
export const checkPMAlertsHttp = functions.https.onRequest(async (req, res) => {
  // Simple CORS
  res.set('Access-Control-Allow-Origin', '*');
  res.set('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.set('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') {
    res.status(204).send('');
    return;
  }

  try {
    const summary = await processUpcomingPMAlerts();
    res.status(200).json({ success: true, summary });
  } catch (error: any) {
    console.error('HTTP trigger error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});
