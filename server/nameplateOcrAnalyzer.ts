import { GoogleGenAI } from '@google/genai';

export interface NameplateOcrInput {
  image_base64: string;
  mime_type?: string;
  equipment_category?: 'transformer' | 'motor' | 'switchgear' | 'cable' | 'general';
}

export interface ExtractedNameplateData {
  equipment_tag: string;
  equipment_name: string;
  equipment_type: string;
  equipment_category?: 'Máy biến áp' | 'Máy cắt' | 'Động cơ' | 'Máy phát' | 'Tủ điện' | 'Inverter';
  manufacturer: string;
  serial_number: string;
  rated_power_kva: number | null;
  rated_power_kw?: number | null;
  primary_voltage_kv: number | null;
  secondary_voltage_v: number | null;
  rated_current_a?: number | null;
  breaking_capacity_ka?: number | null;
  rated_frequency_hz: number | null;
  phases: number | null;
  vector_group: string;
  oil_type: string;
  oil_volume_liters: number | null;
  oil_weight_kg: number | null;
  impedance_percent_z: number | null;
  manufacturing_year: string;
  standards: string;
  ambient_temperature_c: number | null;
  full_nameplate_text: string;
  suggested_field_updates: {
    project_name?: string;
    transformer_tag?: string;
    transformer_type?: string;
    nameplate_match?: boolean;
    pri_test_voltage_kv?: number;
    sec_test_voltage_kv?: number;
    oil_category?: string;
    breaker_tag?: string;
    breaker_rating?: string;
    motor_tag?: string;
    motor_power_kw?: number;
    generator_tag?: string;
    generator_power_kva?: number;
    switchgear_tag?: string;
  };
}

export async function extractNameplateDataWithAi(
  input: NameplateOcrInput
): Promise<{ success: boolean; data?: ExtractedNameplateData; error?: string }> {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return {
      success: false,
      error: 'GEMINI_API_KEY chưa được cấu hình trên hệ thống server.',
    };
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const mimeType = input.mime_type || 'image/jpeg';
    // Remove data URL prefix if present
    let cleanBase64 = input.image_base64;
    if (cleanBase64.includes(';base64,')) {
      cleanBase64 = cleanBase64.split(';base64,')[1];
    }

    const prompt = `
Bạn là Trợ lý AI Chuyên gia Nhận dạng Kỹ thuật Điện cao thế (Optical Character Recognition - OCR High-Voltage Nameplate).
Nhiệm vụ của bạn là đọc và phân tích kỹ lưỡng ảnh chụp nhãn mác kim loại (Nameplate Plate) của thiết bị điện công nghiệp:
- Máy biến áp (Ngâm dầu, Khô Cast Resin)
- Máy cắt (ACB, VCB, SF6)
- Động cơ điện (AC Induction Motor, DC Motor)
- Máy phát điện (Engine Generator)
- Tủ điện phân phối & Đóng cắt (Switchgear, RMU, MSB)

Hãy trích xuất chính xác các thông số kỹ thuật và trả về KẾT QUẢ ĐỊNH DẠNG JSON DUY NHẤT (không dùng markdown code blocks ngoài JSON, hoặc dùng \`\`\`json):
{
  "equipment_tag": "Mã thiết bị gợi ý hoặc đọc được trên nhãn, ví dụ TR-OIL-2000KVA-01, ACB-3200A-01, MTR-PUMP-450KW, GEN-2000KVA, SWG-24KV",
  "equipment_name": "Tên thiết bị đầy đủ, ví dụ: Máy biến áp ngâm dầu 2000kVA 22/0.4kV hoặc Máy cắt ACB 3200A 65kA",
  "equipment_type": "Chủng loại chi tiết",
  "equipment_category": "Máy biến áp HOẶC Máy cắt HOẶC Động cơ HOẶC Máy phát HOẶC Tủ điện",
  "manufacturer": "Hãng sản xuất (ví dụ: ABB, Schneider, Siemens, Dong Anh EEMC, Thibidi, Cummins, Mitsubishi, Stamford, Toshiba, GE)",
  "serial_number": "Số chế tạo / Serial No",
  "rated_power_kva": 2000,
  "rated_power_kw": 1600,
  "primary_voltage_kv": 22.0,
  "secondary_voltage_v": 400,
  "rated_current_a": 52.5,
  "breaking_capacity_ka": 65,
  "rated_frequency_hz": 50,
  "phases": 3,
  "vector_group": "Dyn11 hoặc YNd11...",
  "oil_type": "Mineral Oil (Dầu khoáng) hoặc Silicone hoặc Ester",
  "oil_volume_liters": 1250,
  "oil_weight_kg": 1100,
  "impedance_percent_z": 5.5,
  "manufacturing_year": "2023",
  "standards": "IEC 60076 / IEC 60947 / TCVN 6306 / ANSI IEEE",
  "ambient_temperature_c": 40,
  "full_nameplate_text": "Tất cả các dòng chữ và số đọc được trên nhãn mác kim loại",
  "suggested_field_updates": {
    "transformer_tag": "Mã tag",
    "breaker_tag": "Mã tag",
    "motor_tag": "Mã tag",
    "generator_tag": "Mã tag",
    "switchgear_tag": "Mã tag",
    "nameplate_match": true
  }
}

Chú ý:
- Nếu ảnh bị mờ, phản quang hoặc một số trường không có, hãy suy luận kỹ thuật hợp lý.
- Đảm bảo trả về JSON hợp lệ.
`;

    const response = await ai.models.generateContent({
      model: 'gemini-3.8-flash',
      contents: [
        {
          inlineData: {
            mimeType: mimeType,
            data: cleanBase64,
          },
        },
        prompt,
      ],
    });

    const text = response.text?.trim() || '';
    // Extract JSON block if wrapped
    let jsonStr = text;
    if (text.includes('```json')) {
      jsonStr = text.split('```json')[1].split('```')[0].trim();
    } else if (text.includes('```')) {
      jsonStr = text.split('```')[1].split('```')[0].trim();
    }

    const parsedData = JSON.parse(jsonStr) as ExtractedNameplateData;
    return {
      success: true,
      data: parsedData,
    };
  } catch (err: any) {
    console.error('Lỗi khi phân tích ảnh Nameplate bằng Gemini:', err);
    return {
      success: false,
      error: 'Không thể nhận diện nhãn mác: ' + (err.message || 'Lỗi xử lý ảnh'),
    };
  }
}
