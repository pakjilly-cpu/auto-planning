'use client';

import { useState } from 'react';
import { Settings, Edit2, Check, X } from 'lucide-react';
import { useAppStore } from '@/lib/store';
import { Vendor } from '@/lib/types';

const capabilityLabels: Record<string, string> = {
  normal: '일반',
  shrink: '수축',
  mixing: '교반',
  highFrequency: '고주파',
};

export default function VendorSettings() {
  const { vendors, updateVendor } = useAppStore();
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editData, setEditData] = useState<Partial<Vendor>>({});

  const startEdit = (vendor: Vendor) => {
    setEditingId(vendor.id);
    setEditData({
      lineCount: vendor.lineCount,
      dailyCapacityPerLine: vendor.dailyCapacityPerLine,
      monthlyTarget: vendor.monthlyTarget,
    });
  };

  const handleSave = (vendor: Vendor) => {
    updateVendor({
      ...vendor,
      ...editData,
    });
    setEditingId(null);
    setEditData({});
  };

  const handleCancel = () => {
    setEditingId(null);
    setEditData({});
  };

  return (
    <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
      <div className="flex items-center gap-2 mb-4">
        <Settings className="w-5 h-5 text-gray-500" />
        <h3 className="text-lg font-semibold">외주처 설정</h3>
      </div>
      
      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-gray-200">
              <th className="text-left py-3 px-2 font-medium text-gray-600">외주처</th>
              <th className="text-center py-3 px-2 font-medium text-gray-600">라인수</th>
              <th className="text-center py-3 px-2 font-medium text-gray-600">라인당 일일생산량</th>
              <th className="text-center py-3 px-2 font-medium text-gray-600">월간 목표</th>
              <th className="text-left py-3 px-2 font-medium text-gray-600">가능 공정</th>
              <th className="text-center py-3 px-2 font-medium text-gray-600">수정</th>
            </tr>
          </thead>
          <tbody>
            {vendors.map((vendor: Vendor) => (
              <tr key={vendor.id} className="border-b border-gray-100 hover:bg-gray-50">
                <td className="py-3 px-2">
                  <span className="font-medium">{vendor.name}</span>
                </td>
                <td className="py-3 px-2 text-center">
                  {editingId === vendor.id ? (
                    <input
                      type="number"
                      value={editData.lineCount || ''}
                      onChange={(e) => setEditData({ ...editData, lineCount: parseInt(e.target.value) || 0 })}
                      className="w-16 px-2 py-1 text-center border border-gray-300 rounded"
                    />
                  ) : (
                    vendor.lineCount
                  )}
                </td>
                <td className="py-3 px-2 text-center">
                  {editingId === vendor.id ? (
                    <input
                      type="number"
                      value={editData.dailyCapacityPerLine || ''}
                      onChange={(e) => setEditData({ ...editData, dailyCapacityPerLine: parseInt(e.target.value) || 0 })}
                      className="w-24 px-2 py-1 text-center border border-gray-300 rounded"
                    />
                  ) : (
                    vendor.dailyCapacityPerLine.toLocaleString()
                  )}
                </td>
                <td className="py-3 px-2 text-center">
                  {editingId === vendor.id ? (
                    <input
                      type="number"
                      value={editData.monthlyTarget || ''}
                      onChange={(e) => setEditData({ ...editData, monthlyTarget: parseInt(e.target.value) || 0 })}
                      className="w-28 px-2 py-1 text-center border border-gray-300 rounded"
                    />
                  ) : (
                    vendor.monthlyTarget ? vendor.monthlyTarget.toLocaleString() : '-'
                  )}
                </td>
                <td className="py-3 px-2">
                  <div className="flex flex-wrap gap-1">
                    {vendor.capabilities.map((cap) => (
                      <span
                        key={cap}
                        className={`
                          px-2 py-0.5 text-xs rounded-full
                          ${cap === 'normal' ? 'bg-gray-100 text-gray-600' : ''}
                          ${cap === 'shrink' ? 'bg-blue-100 text-blue-600' : ''}
                          ${cap === 'mixing' ? 'bg-green-100 text-green-600' : ''}
                          ${cap === 'highFrequency' ? 'bg-purple-100 text-purple-600' : ''}
                        `}
                      >
                        {capabilityLabels[cap]}
                      </span>
                    ))}
                  </div>
                </td>
                <td className="py-3 px-2 text-center">
                  {editingId === vendor.id ? (
                    <div className="flex items-center justify-center gap-1">
                      <button
                        onClick={() => handleSave(vendor)}
                        className="p-1 text-green-600 hover:bg-green-100 rounded"
                      >
                        <Check className="w-4 h-4" />
                      </button>
                      <button
                        onClick={handleCancel}
                        className="p-1 text-red-600 hover:bg-red-100 rounded"
                      >
                        <X className="w-4 h-4" />
                      </button>
                    </div>
                  ) : (
                    <button
                      onClick={() => startEdit(vendor)}
                      className="p-1 text-gray-500 hover:bg-gray-200 rounded"
                    >
                      <Edit2 className="w-4 h-4" />
                    </button>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      
      <div className="mt-4 p-3 bg-gray-50 rounded-lg text-xs text-gray-500">
        <p><strong>공정 설명:</strong></p>
        <ul className="mt-1 space-y-0.5">
          <li>• <span className="text-blue-600">수축</span>: 다미, 위드맘, 리니어 가능</li>
          <li>• <span className="text-green-600">교반</span>: 케이코스텍, 리니어, 이시스 가능</li>
          <li>• <span className="text-purple-600">고주파</span>: 위드맘, 리니어 가능</li>
        </ul>
      </div>
    </div>
  );
}
