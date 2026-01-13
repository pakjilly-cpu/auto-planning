'use client';

import { useState } from 'react';
import { Plus, Trash2, Edit2, Check, X } from 'lucide-react';
import { useAppStore } from '@/lib/store';
import { vendorNameMap } from '@/data/defaults';

export default function ClientMapping() {
  const { clientMappings, addClientMapping, removeClientMapping, vendors } = useAppStore();
  
  const [isAdding, setIsAdding] = useState(false);
  const [newCode, setNewCode] = useState('');
  const [newVendor, setNewVendor] = useState('');
  const [editingCode, setEditingCode] = useState<string | null>(null);
  const [editVendor, setEditVendor] = useState('');

  const handleAdd = () => {
    if (newCode.trim() && newVendor) {
      addClientMapping({
        clientCode: newCode.trim().toUpperCase(),
        vendorId: newVendor,
        priority: 1,
      });
      setNewCode('');
      setNewVendor('');
      setIsAdding(false);
    }
  };

  const handleEdit = (code: string) => {
    if (editVendor) {
      addClientMapping({
        clientCode: code,
        vendorId: editVendor,
        priority: 1,
      });
      setEditingCode(null);
      setEditVendor('');
    }
  };

  const startEdit = (code: string, vendorId: string) => {
    setEditingCode(code);
    setEditVendor(vendorId);
  };

  return (
    <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
      <div className="flex items-center justify-between mb-4">
        <h3 className="text-lg font-semibold">고객사-외주처 매칭</h3>
        <button
          onClick={() => setIsAdding(true)}
          className="flex items-center gap-1 text-sm text-blue-600 hover:text-blue-700"
        >
          <Plus className="w-4 h-4" />
          추가
        </button>
      </div>
      
      <div className="space-y-2">
        {/* 새로 추가 */}
        {isAdding && (
          <div className="flex items-center gap-2 p-3 bg-blue-50 rounded-lg">
            <input
              type="text"
              value={newCode}
              onChange={(e) => setNewCode(e.target.value.toUpperCase())}
              placeholder="고객사 코드 (예: CLO)"
              className="flex-1 px-3 py-1.5 text-sm border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
              maxLength={3}
            />
            <select
              value={newVendor}
              onChange={(e) => setNewVendor(e.target.value)}
              className="flex-1 px-3 py-1.5 text-sm border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <option value="">외주처 선택</option>
              {vendors.map((v) => (
                <option key={v.id} value={v.id}>{v.name}</option>
              ))}
            </select>
            <button
              onClick={handleAdd}
              className="p-1.5 text-green-600 hover:bg-green-100 rounded"
            >
              <Check className="w-4 h-4" />
            </button>
            <button
              onClick={() => { setIsAdding(false); setNewCode(''); setNewVendor(''); }}
              className="p-1.5 text-red-600 hover:bg-red-100 rounded"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        )}
        
        {/* 기존 매칭 목록 */}
        {clientMappings.map((mapping) => (
          <div
            key={mapping.clientCode}
            className="flex items-center justify-between p-3 bg-gray-50 rounded-lg"
          >
            {editingCode === mapping.clientCode ? (
              <>
                <span className="font-mono font-medium text-gray-700 w-16">
                  {mapping.clientCode}
                </span>
                <select
                  value={editVendor}
                  onChange={(e) => setEditVendor(e.target.value)}
                  className="flex-1 mx-2 px-3 py-1.5 text-sm border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  {vendors.map((v) => (
                    <option key={v.id} value={v.id}>{v.name}</option>
                  ))}
                </select>
                <div className="flex items-center gap-1">
                  <button
                    onClick={() => handleEdit(mapping.clientCode)}
                    className="p-1.5 text-green-600 hover:bg-green-100 rounded"
                  >
                    <Check className="w-4 h-4" />
                  </button>
                  <button
                    onClick={() => { setEditingCode(null); setEditVendor(''); }}
                    className="p-1.5 text-red-600 hover:bg-red-100 rounded"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>
              </>
            ) : (
              <>
                <div className="flex items-center gap-3">
                  <span className="font-mono font-medium text-gray-700 bg-white px-2 py-1 rounded border">
                    {mapping.clientCode}
                  </span>
                  <span className="text-gray-400">→</span>
                  <span className="font-medium text-gray-700">
                    {vendorNameMap[mapping.vendorId] || mapping.vendorId}
                  </span>
                </div>
                <div className="flex items-center gap-1">
                  <button
                    onClick={() => startEdit(mapping.clientCode, mapping.vendorId)}
                    className="p-1.5 text-gray-500 hover:bg-gray-200 rounded"
                  >
                    <Edit2 className="w-4 h-4" />
                  </button>
                  <button
                    onClick={() => removeClientMapping(mapping.clientCode)}
                    className="p-1.5 text-red-500 hover:bg-red-100 rounded"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              </>
            )}
          </div>
        ))}
        
        {clientMappings.length === 0 && !isAdding && (
          <p className="text-center text-gray-500 py-4">
            등록된 매칭이 없습니다.
          </p>
        )}
      </div>
      
      <p className="mt-4 text-xs text-gray-500">
        * 제품코드의 영문 3자리 (예: 9CLO... → CLO)가 해당 외주처에 우선 배정됩니다.
      </p>
    </div>
  );
}
