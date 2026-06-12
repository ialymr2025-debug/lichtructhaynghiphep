import React, { useState, useEffect } from 'react';

interface SignatureManagerProps {
  staffList: string[];
  signatures: Record<string, string>;
  onSignaturesChange: (signatures: Record<string, string>) => void;
}

export default function SignatureManager({ staffList, signatures, onSignaturesChange }: SignatureManagerProps) {
  const [saving, setSaving] = useState<string | null>(null);

  const handleUpload = async (name: string, file: File) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      const base64 = e.target?.result as string;
      setSaving(name);
      try {
        const res = await fetch('/api/signatures', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name, data: base64 })
        });
        if (res.ok) {
          const newSigs = { ...signatures, [name]: base64 };
          onSignaturesChange(newSigs);
        }
      } catch (e) {
        console.error("Failed to save signature", e);
      } finally {
        setSaving(null);
      }
    };
    reader.readAsDataURL(file);
  };

  return (
    <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
      <h3 className="text-lg font-bold text-gray-800 mb-4 flex items-center gap-2">
        <span className="text-blue-600">✍️</span> Quản lý chữ ký (65+ người)
      </h3>
      <p className="text-sm text-gray-500 mb-4 italic">
        * Tải lên ảnh chữ ký cho từng người. Ảnh sẽ được lưu vĩnh viễn trên hệ thống.
      </p>
      
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 max-h-[500px] overflow-y-auto pr-2">
        {staffList.map(name => (
          <div key={name} className="p-3 border rounded-lg flex flex-col gap-2 bg-gray-50">
            <div className="flex justify-between items-center">
              <span className="font-medium text-gray-700">{name}</span>
              {signatures[name] ? (
                <span className="text-xs bg-green-100 text-green-700 px-2 py-0.5 rounded-full">Đã có</span>
              ) : (
                <span className="text-xs bg-red-100 text-red-700 px-2 py-0.5 rounded-full">Chưa có</span>
              )}
            </div>
            
            {signatures[name] && (
              <div className="h-16 bg-white border rounded flex items-center justify-center overflow-hidden">
                <img src={signatures[name]} alt="signature" className="max-h-full object-contain" />
              </div>
            )}

            <label className={`
              cursor-pointer text-center py-1.5 px-3 rounded text-sm font-medium transition-colors
              ${saving === name ? 'bg-gray-300 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700 text-white'}
            `}>
              {saving === name ? 'Đang lưu...' : (signatures[name] ? 'Thay đổi ảnh' : 'Tải ảnh lên')}
              <input 
                type="file" 
                className="hidden" 
                accept="image/*" 
                onChange={(e) => e.target.files?.[0] && handleUpload(name, e.target.files[0])}
                disabled={saving === name}
              />
            </label>
          </div>
        ))}
      </div>
    </div>
  );
}
