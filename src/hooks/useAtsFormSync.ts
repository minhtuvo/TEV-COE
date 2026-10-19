import React, { useState, useEffect, useRef } from 'react';
import { getAtsPayload, saveAtsPayload } from '../services/netaAtsDataService';

function mergeWithDefault<T extends Record<string, any>>(defaultData: T, cached: any): T {
  if (!cached || typeof cached !== 'object' || Object.keys(cached).length === 0) {
    return defaultData;
  }
  const merged: any = {
    ...defaultData,
    ...cached,
  };
  if (defaultData.site_info || cached.site_info) {
    merged.site_info = {
      ...(defaultData.site_info || {}),
      ...(cached.site_info || {}),
    };
  }
  if (defaultData.station_info || cached.station_info) {
    merged.station_info = {
      ...(defaultData.station_info || {}),
      ...(cached.station_info || {}),
    };
  }
  if (defaultData.visual_inspection && cached.visual_inspection) {
    merged.visual_inspection = {
      ...defaultData.visual_inspection,
      ...cached.visual_inspection,
    };
  }
  if (defaultData.electrical_tests && cached.electrical_tests) {
    merged.electrical_tests = {
      ...defaultData.electrical_tests,
      ...cached.electrical_tests,
    };
  }
  return merged as T;
}

/**
 * Universal hook for synchronizing checklist form state with localStorage
 * and the TEV Standard 4-part AI Report.
 */
export function useAtsFormSync<T extends Record<string, any>>(
  specId: string,
  defaultData: T
): [T, React.Dispatch<React.SetStateAction<T>>] {
  const [data, setData] = useState<T>(() => {
    const cached = getAtsPayload(specId);
    return mergeWithDefault(defaultData, cached);
  });

  const lastSavedJsonRef = useRef<string>(
    JSON.stringify(mergeWithDefault(defaultData, getAtsPayload(specId)))
  );

  // When specId changes, reload appropriate data
  useEffect(() => {
    const cached = getAtsPayload(specId);
    const merged = mergeWithDefault(defaultData, cached);
    lastSavedJsonRef.current = JSON.stringify(merged);
    setData(merged);
  }, [specId]);

  // Synchronize changes to storage and trigger broadcast cleanly after render
  useEffect(() => {
    if (data) {
      const json = JSON.stringify(data);
      if (json !== lastSavedJsonRef.current) {
        lastSavedJsonRef.current = json;
        saveAtsPayload(specId, data, 'form');
      }
    }
  }, [specId, data]);

  // Listen for reset or external updates (e.g. from Project Manager, JSON editor, or scenario resets)
  useEffect(() => {
    const handleUpdate = (e: any) => {
      if (e.detail?.specId === specId) {
        if (e.detail.source === 'form') {
          // Ignore reflection of our own updates
          return;
        }
        if (e.detail.reset) {
          lastSavedJsonRef.current = JSON.stringify(defaultData);
          setData(defaultData);
        } else if (e.detail.payload) {
          const merged = mergeWithDefault(defaultData, e.detail.payload);
          const newJson = JSON.stringify(merged);
          if (newJson !== lastSavedJsonRef.current) {
            lastSavedJsonRef.current = newJson;
            setData(merged);
          }
        }
      }
    };
    window.addEventListener('tev_ats_payload_updated', handleUpdate);
    return () => window.removeEventListener('tev_ats_payload_updated', handleUpdate);
  }, [specId, defaultData]);

  return [data, setData];
}

