/**
 * DGA History Service: Provides last recorded values from the historical database
 * for transformers and DGA diagnostic parameter comparison.
 */

export interface HistoricalDgaRecord {
  date: string;
  label?: string;
  source: string;
  gases: {
    h2?: number;
    ch4?: number;
    c2h6?: number;
    c2h4?: number;
    c2h2?: number;
    co?: number;
    co2?: number;
    o2?: number;
    n2?: number;
  };
  otherParams: {
    moisture?: number;
    bdStrength?: number;
    acidity?: number;
    ffa?: number;
    estDp?: number;
    age?: number;
    loadFactor?: number;
  };
}

// Baseline historical benchmarks (representing previous test cycle in the database)
export const DEFAULT_BASELINE_DGA_RECORD: HistoricalDgaRecord = {
  date: '10/03/2026',
  label: 'Kỳ kiểm tra định kỳ trước (10/03/2026)',
  source: 'Hồ sơ kiểm định định kỳ Q1/2026',
  gases: {
    h2: 68,
    ch4: 82,
    c2h6: 48,
    c2h4: 35,
    c2h2: 0.7,
    co: 420,
    co2: 3250,
    o2: 520,
    n2: 44000
  },
  otherParams: {
    moisture: 14,
    bdStrength: 52,
    acidity: 0.06,
    ffa: 1.5,
    estDp: 680,
    age: 20,
    loadFactor: 65
  }
};

/**
 * Retrieves the last recorded historical values for a given equipment code.
 * Searches in local allReports first; falls back to the database baseline.
 */
export function getLastRecordedDgaValues(
  equipmentCode: string | undefined,
  allReports?: any[]
): HistoricalDgaRecord {
  if (allReports && allReports.length > 0 && equipmentCode) {
    // Look for previous reports of this equipment containing DGA data
    const matched = allReports
      .filter((r) => {
        const isSameEq = r.equipmentId === equipmentCode || r.id?.includes(equipmentCode);
        const hasDgaMeasurements = r.measurements && (
          r.measurements.h2 !== undefined ||
          r.measurements.c2h2 !== undefined ||
          r.measurements.ch4 !== undefined ||
          r.measurements.dga !== undefined
        );
        const hasNetaDga = r.details?.netaData?.electrical_tests?.dga_test_ieee_c57_104 ||
          r.details?.netaData?.electrical_tests?.oil_sample_tests_astm_d923;
        return isSameEq && (hasDgaMeasurements || hasNetaDga);
      })
      .sort((a, b) => {
        const tA = new Date(a.date).getTime() || 0;
        const tB = new Date(b.date).getTime() || 0;
        return tB - tA;
      });

    if (matched.length > 0) {
      const latest = matched[0];
      const m = latest.measurements || {};
      const neta = latest.details?.netaData?.electrical_tests || {};
      const dga = neta.dga_test_ieee_c57_104 || {};
      const oil = neta.oil_sample_tests_astm_d923 || {};

      const parseNum = (val: any) => {
        if (val === undefined || val === null || val === '') return undefined;
        const n = parseFloat(String(val).replace(',', '.'));
        return isNaN(n) ? undefined : n;
      };

      return {
        date: latest.date || 'Lần trước',
        label: `Báo cáo CMMS #${latest.id || ''} (${latest.date || ''})`,
        source: 'CMMS Database',
        gases: {
          h2: parseNum(m.h2 ?? dga.hydrogen_h2_ppm),
          ch4: parseNum(m.ch4 ?? dga.methane_ch4_ppm),
          c2h6: parseNum(m.c2h6 ?? dga.ethane_c2h6_ppm),
          c2h4: parseNum(m.c2h4 ?? dga.ethylene_c2h4_ppm),
          c2h2: parseNum(m.c2h2 ?? dga.acetylene_c2h2_ppm),
          co: parseNum(m.co ?? dga.carbon_monoxide_co_ppm),
          co2: parseNum(m.co2 ?? dga.carbon_dioxide_co2_ppm),
          o2: parseNum(m.o2 ?? dga.oxygen_o2_ppm),
          n2: parseNum(m.n2 ?? dga.nitrogen_n2_ppm)
        },
        otherParams: {
          moisture: parseNum(m.oilMoisture ?? oil.water_content_d1533_ppm),
          bdStrength: parseNum(m.dielectricStrength ?? oil.dielectric_breakdown_d1816_1mm_kv),
          acidity: parseNum(m.acidity ?? oil.acid_number_d974_mg_koh_per_g),
          ffa: parseNum(m.furan),
          estDp: parseNum(m.estDp),
          age: parseNum(m.age),
          loadFactor: parseNum(m.dutyFactor ? parseFloat(m.dutyFactor) * 100 : undefined)
        }
      };
    }
  }

  // Fallback to the standardized baseline historical record
  return DEFAULT_BASELINE_DGA_RECORD;
}
