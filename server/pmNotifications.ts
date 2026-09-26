import nodemailer from 'nodemailer';

export interface PMWorkOrder {
  id: string;
  title: string;
  description?: string;
  equipmentId?: string | string[];
  customerId?: string;
  factory?: string;
  assignedTo?: string;
  status: string;
  priority?: string;
  type: string;
  dueDate?: string;
  pmFrequency?: string;
  notified3DaysBefore?: boolean;
  lastNotificationSentAt?: string;
  lastNotificationStatus?: string;
  [key: string]: any;
}

export interface EquipmentItem {
  id: string;
  name: string;
  customer?: string;
  customerId?: string;
  location?: string;
  status?: string;
  lastInspection?: string;
  nextInspection?: string;
  [key: string]: any;
}

export interface UpcomingPMTask {
  workOrderId: string;
  title: string;
  equipmentId: string;
  equipmentName: string;
  factory?: string;
  dueDate: string;
  daysRemaining: number;
  priority: string;
  pmFrequency?: string;
  assignedTo?: string;
  recipients: string[];
  alreadyNotified: boolean;
  isExact3Days: boolean;
}

// Generate styled HTML email for PM 3-day reminder
export function generatePMEmailHtml(data: {
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

  const priorityLabel =
    data.priority === 'urgent'
      ? 'Khẩn cấp'
      : data.priority === 'high'
      ? 'Cao'
      : data.priority === 'medium'
      ? 'Trung bình'
      : 'Thấp';

  return `
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Cảnh báo bảo trì định kỳ PM</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #f1f5f9; color: #1e293b; margin: 0; padding: 24px 12px; }
    .wrapper { max-width: 620px; margin: 0 auto; background: #ffffff; border-radius: 14px; overflow: hidden; border: 1px solid #e2e8f0; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.05); }
    .header { background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); color: #ffffff; padding: 28px 24px; text-align: left; }
    .brand { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; color: #38bdf8; margin-bottom: 8px; }
    .title { margin: 0; font-size: 21px; font-weight: 700; color: #ffffff; line-height: 1.3; }
    .body { padding: 28px 24px; }
    .badge-pill { display: inline-flex; align-items: center; padding: 6px 14px; font-size: 13px; font-weight: 700; border-radius: 9999px; background-color: #fef3c7; color: #92400e; border: 1px solid #fde68a; margin-bottom: 20px; }
    .details-table { width: 100%; border-collapse: separate; border-spacing: 0; margin-top: 16px; margin-bottom: 24px; border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden; }
    .details-table tr:not(:last-child) td { border-bottom: 1px solid #e2e8f0; }
    .details-table td { padding: 12px 16px; font-size: 14px; }
    .details-table td.label { background-color: #f8fafc; font-weight: 600; color: #64748b; width: 38%; }
    .details-table td.val { color: #0f172a; font-weight: 500; }
    .instructions { background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 16px; margin-bottom: 24px; font-size: 13px; line-height: 1.6; color: #166534; }
    .footer { padding: 20px 24px; background: #f8fafc; border-top: 1px solid #e2e8f0; font-size: 12px; color: #64748b; text-align: center; line-height: 1.5; }
  </style>
</head>
<body>
  <div class="wrapper">
    <div class="header">
      <div class="brand">Renewable CMMS • Automated PM Notification</div>
      <h1 class="title">🚨 Cảnh Báo: Lịch Bảo Trì Định Kỳ (PM) Còn 3 Ngày</h1>
    </div>

    <div class="body">
      <div class="badge-pill">
        ⏰ Sắp đến hạn: Còn ${data.daysRemaining} ngày (Hạn: ${data.dueDateStr})
      </div>

      <p style="font-size: 15px; line-height: 1.6; color: #334155; margin-top: 0;">
        Kính gửi Quý Kỹ sư & Ban Quản trị Vận hành,<br/>
        Hệ thống CMMS ghi nhận phiếu công việc bảo trì định kỳ sau đây sẽ đến hạn trong <strong>3 ngày tới</strong>. Xin vui lòng kiểm tra kế hoạch nhân lực, phiếu thao tác an toàn và vật tư thay thế tương ứng:
      </p>

      <table class="details-table">
        <tr>
          <td class="label">Phiếu công việc (PM):</td>
          <td class="val"><strong>${data.title}</strong></td>
        </tr>
        <tr>
          <td class="label">Mã phiếu (ID):</td>
          <td class="val"><code style="background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-family: monospace; font-size: 13px;">${data.orderId}</code></td>
        </tr>
        <tr>
          <td class="label">Thiết bị bảo trì:</td>
          <td class="val"><strong>${data.equipmentName}</strong> (${data.equipmentId})</td>
        </tr>
        ${data.factory ? `
        <tr>
          <td class="label">Nhà máy / Vị trí:</td>
          <td class="val">${data.factory}</td>
        </tr>` : ''}
        <tr>
          <td class="label">Ngày đến hạn:</td>
          <td class="val" style="color: #b45309; font-weight: 700; font-size: 15px;">${data.dueDateStr}</td>
        </tr>
        <tr>
          <td class="label">Mức độ ưu tiên:</td>
          <td class="val"><span style="color: ${priorityColor}; font-weight: 700; text-transform: uppercase;">${priorityLabel}</span></td>
        </tr>
        ${data.pmFrequency ? `
        <tr>
          <td class="label">Chu kỳ bảo trì:</td>
          <td class="val">${data.pmFrequency}</td>
        </tr>` : ''}
        ${data.assignedTo ? `
        <tr>
          <td class="label">Kỹ thuật viên phụ trách:</td>
          <td class="val">${data.assignedTo}</td>
        </tr>` : ''}
        ${data.description ? `
        <tr>
          <td class="label">Nội dung công việc:</td>
          <td class="val">${data.description}</td>
        </tr>` : ''}
      </table>

      <div class="instructions">
        <strong>📋 Checklist chuẩn bị trước khi bảo trì:</strong><br/>
        1. Liên hệ trung tâm điều độ / chủ đầu tư để đăng ký lịch cắt điện (nếu cần cô lập thiết bị).<br/>
        2. Soát xét tồn kho phụ tùng, dầu cách điện và dụng cụ thử nghiệm chuyên dụng.<br/>
        3. Phổ biến biện pháp an toàn lao động và cấp phiếu công tác (Work Permit).
      </div>
    </div>

    <div class="footer">
      Thông báo tự động từ Hệ thống Renewable CMMS & Firebase Cloud Functions.<br/>
      Email được gửi tự động 3 ngày trước ngày thực hiện PM theo lịch định kỳ.
    </div>
  </div>
</body>
</html>
`;
}

// Helper to get Nodemailer transporter
export function getEmailTransporter() {
  const host = process.env.SMTP_HOST;
  const port = process.env.SMTP_PORT ? parseInt(process.env.SMTP_PORT, 10) : 587;
  const user = process.env.SMTP_USER;
  const pass = process.env.SMTP_PASS;

  if (host && user && pass) {
    return nodemailer.createTransport({
      host,
      port,
      secure: port === 465,
      auth: { user, pass },
    });
  }

  return null;
}

// Find all PM work orders that are due in 3 days (or within 3 days and not yet completed)
export function findUpcomingPMOrders(
  workOrders: PMWorkOrder[],
  allEquipment: EquipmentItem[],
  targetDaysAhead: number = 3
): UpcomingPMTask[] {
  const now = new Date();
  // Target date at 00:00
  const targetDate = new Date();
  targetDate.setDate(now.getDate() + targetDaysAhead);
  const targetDateStr = targetDate.toISOString().split('T')[0];

  const results: UpcomingPMTask[] = [];

  for (const order of workOrders) {
    // Only preventive maintenance work orders
    if (order.type !== 'preventive') continue;
    // Skip completed or cancelled
    if (order.status === 'completed' || order.status === 'cancelled') continue;
    if (!order.dueDate) continue;

    const dueDateObj = new Date(order.dueDate);
    if (isNaN(dueDateObj.getTime())) continue;

    const dueDateStr = dueDateObj.toISOString().split('T')[0];
    
    // Difference in whole days
    const diffTime = dueDateObj.getTime() - now.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    // Target condition: task is due in exactly targetDaysAhead (e.g. 3 days),
    // or within 1 to 3 days if not yet completed and not yet notified
    const isExact3Days = dueDateStr === targetDateStr || diffDays === targetDaysAhead;
    const isUpcomingInWindow = diffDays >= 0 && diffDays <= targetDaysAhead;

    if (isUpcomingInWindow) {
      // Find equipment name
      let eqName = 'Thiết bị';
      const rawEqId = Array.isArray(order.equipmentId) ? order.equipmentId[0] : order.equipmentId;
      if (rawEqId) {
        const found = allEquipment.find(e => e.id === rawEqId);
        if (found) {
          eqName = found.name;
        }
      }

      // Collect recipient candidates
      const recipients = new Set<string>();
      const defaultEmail = process.env.DEFAULT_ALERT_EMAIL || 'sgm1707@gmail.com';
      recipients.add(defaultEmail);

      if (order.assignedTo && order.assignedTo.includes('@')) {
        recipients.add(order.assignedTo);
      }

      results.push({
        workOrderId: order.id,
        title: order.title || `Bảo trì định kỳ - ${eqName}`,
        equipmentId: rawEqId || 'N/A',
        equipmentName: eqName,
        factory: order.factory,
        dueDate: dueDateStr,
        daysRemaining: diffDays,
        priority: order.priority || 'medium',
        pmFrequency: order.pmFrequency,
        assignedTo: order.assignedTo,
        recipients: Array.from(recipients),
        alreadyNotified: !!order.notified3DaysBefore,
        isExact3Days,
      });
    }
  }

  return results;
}

// Send PM notification emails for the identified tasks
export async function sendPMNotifications(
  tasks: UpcomingPMTask[],
  options?: {
    forceSend?: boolean; // Send even if alreadyNotified is true (e.g. manual test)
    customRecipient?: string;
  }
) {
  const transporter = getEmailTransporter();
  const results = [];
  const fromAddress = process.env.NOTIFICATION_EMAIL_FROM || '"Renewable CMMS" <notifications@cmms-system.com>';

  for (const task of tasks) {
    if (task.alreadyNotified && !options?.forceSend) {
      results.push({
        workOrderId: task.workOrderId,
        title: task.title,
        status: 'skipped',
        reason: 'Already notified previously',
      });
      continue;
    }

    const recipients = options?.customRecipient ? [options.customRecipient] : task.recipients;
    const emailHtml = generatePMEmailHtml({
      orderId: task.workOrderId,
      title: task.title,
      equipmentName: task.equipmentName,
      equipmentId: task.equipmentId,
      factory: task.factory,
      dueDateStr: new Date(task.dueDate).toLocaleDateString('vi-VN'),
      daysRemaining: task.daysRemaining,
      priority: task.priority,
      assignedTo: task.assignedTo,
      pmFrequency: task.pmFrequency,
    });

    const subject = `[CẢNH BÁO PM - CÒN ${task.daysRemaining} NGÀY] ${task.title} - ${task.equipmentName} (Hạn: ${new Date(task.dueDate).toLocaleDateString('vi-VN')})`;

    if (transporter) {
      try {
        const info = await transporter.sendMail({
          from: fromAddress,
          to: recipients.join(', '),
          subject,
          html: emailHtml,
        });

        results.push({
          workOrderId: task.workOrderId,
          title: task.title,
          status: 'sent',
          recipients,
          messageId: info.messageId,
          timestamp: new Date().toISOString(),
        });
      } catch (err: any) {
        console.error(`Error sending email to ${recipients.join(', ')}:`, err);
        results.push({
          workOrderId: task.workOrderId,
          title: task.title,
          status: 'failed',
          recipients,
          error: err.message,
          timestamp: new Date().toISOString(),
        });
      }
    } else {
      // SMTP not configured, operate in simulated/preview mode
      console.log(`[SIMULATED EMAIL] PM 3-Day Alert for Order ${task.workOrderId}`);
      console.log(`Subject: ${subject}`);
      console.log(`Recipients: ${recipients.join(', ')}`);
      
      results.push({
        workOrderId: task.workOrderId,
        title: task.title,
        status: 'simulated',
        recipients,
        subject,
        htmlPreview: emailHtml,
        notice: 'Email được mô phỏng thành công. Để gửi email thực tế đến hộp thư, cấu hình SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS trong Settings -> Secrets.',
        timestamp: new Date().toISOString(),
      });
    }
  }

  return {
    totalTasks: tasks.length,
    processedCount: results.length,
    smtpConfigured: !!transporter,
    results,
  };
}
