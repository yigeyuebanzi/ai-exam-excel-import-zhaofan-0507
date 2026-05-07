'use client';

import React, { useState, useEffect, useMemo } from 'react';
import { Upload, FileSpreadsheet, AlertCircle, CheckCircle2, Loader2, Save, History, Search, Download, Trash2, Plus } from 'lucide-react';
import * as XLSX from 'xlsx';

// --- 1. 配置与定义 ---
const FIELD_DEFINITIONS = {
  externalCode: { label: '外部编码', required: false, type: 'string' },
  senderName: { label: '发件人姓名', required: true, type: 'string' },
  senderPhone: { label: '发件人电话', required: true, type: 'string' },
  senderAddress: { label: '发件人地址', required: true, type: 'string' },
  receiverName: { label: '收件人姓名', required: true, type: 'string' },
  receiverPhone: { label: '收件人电话', required: true, type: 'string' },
  receiverAddress: { label: '收件人地址', required: true, type: 'string' },
  weight: { label: '重量 (kg)', required: true, type: 'number', positive: true },
  quantity: { label: '件数', required: true, type: 'number', positiveInteger: true },
  tempZone: { label: '温层', required: true, type: 'enum', options: ['常温', '冷藏', '冷冻'] },
  remark: { label: '备注', required: false, type: 'string' },
} as const;

type OrderField = keyof typeof FIELD_DEFINITIONS;

const ALIAS_MAP: Record<string, OrderField> = {
  // 外部编码
  '外部编码': 'externalCode', '外部单号': 'externalCode', '外部订单号': 'externalCode',
  'Ref Code': 'externalCode', 'ref code': 'externalCode',
  '客户单号': 'externalCode', '订单编号': 'externalCode', '单号': 'externalCode',
  // 发件人
  '发件人姓名': 'senderName', '发件人': 'senderName', '寄件人': 'senderName', 'Sender': 'senderName', 'sender': 'senderName',
  '寄件人姓名': 'senderName', '寄件人名称': 'senderName',
  '发货人': 'senderName', '发货人姓名': 'senderName', // 电商模板
  // 发件电话
  '发件人电话': 'senderPhone', '发件电话': 'senderPhone', '寄件人电话': 'senderPhone', 'Sender Tel': 'senderPhone',
  'sender tel': 'senderPhone', '寄件电话': 'senderPhone',
  '发货电话': 'senderPhone', '发货人电话': 'senderPhone', // 电商模板
  // 发件地址
  '发件人地址': 'senderAddress', '发件地址': 'senderAddress', '寄件人地址': 'senderAddress', 'Sender Address': 'senderAddress',
  'sender address': 'senderAddress', '寄件地址': 'senderAddress',
  '发货地址': 'senderAddress', '发货人地址': 'senderAddress', // 电商模板
  // 收件人
  '收件人姓名': 'receiverName', '收件人': 'receiverName', '收货人': 'receiverName', '收方': 'receiverName',
  'Receiver': 'receiverName', 'receiver': 'receiverName', '收货人姓名': 'receiverName', '收件人名称': 'receiverName',
  // 收件电话
  '收件人电话': 'receiverPhone', '收件电话': 'receiverPhone', '收货人电话': 'receiverPhone', 'Receiver Tel': 'receiverPhone',
  'receiver tel': 'receiverPhone', '收货电话': 'receiverPhone',
  // 收件地址
  '收件人地址': 'receiverAddress', '收件地址': 'receiverAddress', '收货人地址': 'receiverAddress', 'Receiver Address': 'receiverAddress',
  'receiver address': 'receiverAddress', '收货地址': 'receiverAddress',
  // 重量
  '重量(kg)': 'weight', '重量 (kg)': 'weight', '重量': 'weight',
  'Weight(kg)': 'weight', 'weight(kg)': 'weight', 'Weight(KG)': 'weight', 'weight(KG)': 'weight',
  'Weight': 'weight', 'weight': 'weight', '货物重量': 'weight',
  // 件数
  '件数': 'quantity', '数量': 'quantity', 'Qty': 'quantity', 'qty': 'quantity', '包裹数量': 'quantity',
  // 温层
  '温层': 'tempZone', '温度要求': 'tempZone', 'Temp Zone': 'tempZone', 'temp zone': 'tempZone',
  '存储条件': 'tempZone', '温度': 'tempZone',
  // 备注
  '备注': 'remark', 'Note': 'remark', 'note': 'remark', '附加说明': 'remark',
  '客户备注': 'remark', // 电商模板
};

// --- 2. 核心逻辑函数 ---
function validateOrder(order: any, allOrders: any[] = []) {
  const errors: { field: string; message: string }[] = [];
  
  Object.entries(FIELD_DEFINITIONS).forEach(([field, def]) => {
    const d = def as any;
    const val = order[field];
    const isPresent = val !== undefined && val !== null && String(val).trim() !== '';

    if (d.required && !isPresent) {
      errors.push({ field, message: `${d.label}是必填项` });
    } else if (isPresent) {
      if (d.type === 'number') {
        const num = Number(val);
        if (isNaN(num)) errors.push({ field, message: `${d.label}必须是数字` });
        else if (d.positive && num <= 0) errors.push({ field, message: `${d.label}必须大于0` });
        else if (d.positiveInteger && (!Number.isInteger(num) || num <= 0)) errors.push({ field, message: `${d.label}必须是正整数` });
      }
      if (d.type === 'enum' && d.options && !d.options.includes(val)) {
        errors.push({ field, message: `${d.label}值非法，可选：${d.options.join('/')}` });
      }
      if (field.toLowerCase().includes('phone') && !/^1[3-9]\d{9}$/.test(String(val).trim())) {
        errors.push({ field, message: `${d.label}格式错误` });
      }
    }
  });

  // 重复检测
  if (order.externalCode) {
    const isDuplicate = allOrders.some(o => o.id !== order.id && o.externalCode === order.externalCode);
    if (isDuplicate) errors.push({ field: 'externalCode', message: '外部编码在该批次中重复' });
  }

  return errors;
}

// --- 3. UI 组件 ---
export default function OrderImporter() {
  const [orders, setOrders] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [view, setView] = useState<'upload' | 'preview' | 'history'>('upload');
  const [history, setHistory] = useState<any[]>([]);
  const [searchTerm, setSearchTerm] = useState('');

  // 表头标准化：去空格、转小写、中文括号转英文
  const normalizeHeader = (raw: string) => {
    return raw
      .replace(/[\s\u00A0]+/g, '')      // 去除所有空格（包括全角空格）
      .replace(/[（）]/g, m => m === '（' ? '(' : ')')  // 中文括号转英文
      .toLowerCase();
  };

  // 构建标准化后的别名映射表（只做一次）
  const NORMALIZED_ALIAS_MAP = useMemo(() => {
    const map: Record<string, OrderField> = {};
    Object.entries(ALIAS_MAP).forEach(([key, val]) => {
      map[normalizeHeader(key)] = val;
    });
    return map;
  }, []);

  // 处理上传
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setProgress(10);

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = evt.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        
        setProgress(40);
        
        // 自动识别表头（遍历所有sheet，使用标准化匹配）
        let headerIdx = -1;
        let mapping: Record<number, OrderField> = {};
        let rows: any[][] = [];
        
        for (const sheetName of workbook.SheetNames) {
          const sheetRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 }) as any[][];
          for (let i = 0; i < Math.min(sheetRows.length, 15); i++) {
            const row = sheetRows[i];
            if (!row) continue;
            const matches = row.filter(c => NORMALIZED_ALIAS_MAP[normalizeHeader(String(c || ''))]);
            if (matches.length >= 3) {
              headerIdx = i;
              rows = sheetRows;
              row.forEach((c, idx) => {
                const field = NORMALIZED_ALIAS_MAP[normalizeHeader(String(c || ''))];
                if (field) mapping[idx] = field;
              });
              break;
            }
          }
          if (headerIdx !== -1) break;
        }

        if (headerIdx === -1) throw new Error('未能识别模板格式，请确保表头包含关键字段');

        const parsedOrders = rows.slice(headerIdx + 1)
          .filter(row => row.some(c => c))
          .map((row, i) => {
            const order: any = { id: `row-${Date.now()}-${i}` };
            Object.entries(mapping).forEach(([colIdx, field]) => {
              order[field] = row[Number(colIdx)];
            });
            return order;
          });

        setProgress(80);
        const validated = parsedOrders.map(o => ({ ...o, errors: validateOrder(o, parsedOrders) }));
        setOrders(validated);
        setView('preview');
        setProgress(100);
      } catch (err: any) {
        alert(err.message);
      } finally {
        setTimeout(() => setLoading(false), 300);
      }
    };
    reader.readAsBinaryString(file);
  };

  const updateCell = (idx: number, field: string, val: any) => {
    const updated = [...orders];
    updated[idx] = { ...updated[idx], [field]: val };
    // 重新校验全量（为了处理重复）
    const reValidated = updated.map(o => ({ ...o, errors: validateOrder(o, updated) }));
    setOrders(reValidated);
  };

  const deleteRow = (idx: number) => {
    const updated = orders.filter((_, i) => i !== idx);
    setOrders(updated.map(o => ({ ...o, errors: validateOrder(o, updated) })));
  };

  const addRow = () => {
    const newRow = { id: `new-${Date.now()}`, errors: [] };
    setOrders([...orders, { ...newRow, errors: validateOrder(newRow, [...orders, newRow]) }]);
  };

  const submitOrders = async () => {
    const hasErrors = orders.some(o => o.errors.length > 0);
    if (hasErrors) return alert('请先修正所有错误行后再提交');

    setLoading(true);
    setProgress(0);
    
    // 模拟提交
    for (let p = 0; p <= 100; p += 20) {
      setProgress(p);
      await new Promise(r => setTimeout(r, 150));
    }

    const batch = { id: Date.now(), time: new Date().toLocaleString(), count: orders.length, data: orders };
    const newHistory = [batch, ...history];
    setHistory(newHistory);
    localStorage.setItem('v_orders', JSON.stringify(newHistory));
    
    alert('提交成功！订单已保存至系统。');
    setOrders([]);
    setView('history');
    setLoading(false);
  };

  useEffect(() => {
    const saved = localStorage.getItem('v_orders');
    if (saved) setHistory(JSON.parse(saved));
  }, []);

  const filteredHistory = useMemo(() => {
    if (!searchTerm) return history;
    return history.filter(h => JSON.stringify(h.data).includes(searchTerm));
  }, [history, searchTerm]);

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-[#1E293B] font-sans">
      {/* 导航栏 */}
      <header className="bg-white border-b sticky top-0 z-30 px-6 py-4 flex items-center justify-between shadow-sm">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg text-white">
            <FileSpreadsheet className="w-6 h-6" />
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tight">万能导入系统</h1>
            <p className="text-xs text-gray-400 font-medium">Smart Excel Import Engine v1.0</p>
          </div>
        </div>
        <div className="flex bg-gray-100 p-1 rounded-xl">
          {(['upload', 'history'] as const).map(t => (
            <button key={t} onClick={() => setView(t)} className={`px-5 py-2 rounded-lg text-sm font-bold transition ${view === t || (view === 'preview' && t === 'upload') ? 'bg-white text-blue-600 shadow-sm' : 'text-gray-500 hover:text-gray-700'}`}>
              {t === 'upload' ? '导入下单' : '历史运单'}
            </button>
          ))}
        </div>
      </header>

      <main className="max-w-[1600px] mx-auto p-6">
        {loading && (
          <div className="fixed inset-0 bg-white/80 backdrop-blur-sm z-50 flex flex-col items-center justify-center">
            <Loader2 className="w-10 h-10 text-blue-600 animate-spin mb-4" />
            <div className="w-64 h-2 bg-gray-100 rounded-full overflow-hidden">
              <div className="bg-blue-600 h-full transition-all duration-300" style={{ width: `${progress}%` }} />
            </div>
            <p className="mt-2 text-sm font-bold text-blue-600">{progress}% 处理中...</p>
          </div>
        )}

        {view === 'upload' && (
          <div className="max-w-2xl mx-auto mt-20">
            <div className="bg-white border-2 border-dashed border-gray-200 rounded-3xl p-16 text-center hover:border-blue-400 transition group relative">
              <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="absolute inset-0 opacity-0 cursor-pointer" />
              <div className="bg-blue-50 w-20 h-20 rounded-2xl flex items-center justify-center mx-auto mb-6 group-hover:scale-110 transition">
                <Upload className="w-10 h-10 text-blue-600" />
              </div>
              <h2 className="text-2xl font-bold mb-2">点击或拖拽上传 Excel</h2>
              <p className="text-gray-400">支持多模板智能匹配，解析 1000+ 条数据仅需数秒</p>
            </div>
          </div>
        )}

        {view === 'preview' && (
          <div className="space-y-4">
            <div className="flex items-center justify-between bg-white p-4 rounded-2xl border shadow-sm">
              <div className="flex gap-4 items-center">
                <div className="flex items-center gap-2 px-4 py-2 bg-blue-50 text-blue-700 rounded-xl font-bold text-sm">
                  <CheckCircle2 className="w-4 h-4" /> 共 {orders.length} 条
                </div>
                {orders.some(o => o.errors.length > 0) && (
                  <div className="flex items-center gap-2 px-4 py-2 bg-red-50 text-red-700 rounded-xl font-bold text-sm">
                    <AlertCircle className="w-4 h-4" /> {orders.filter(o => o.errors.length > 0).length} 条异常
                  </div>
                )}
                <button onClick={addRow} className="flex items-center gap-2 px-4 py-2 border hover:bg-gray-50 rounded-xl font-bold text-sm">
                  <Plus className="w-4 h-4" /> 新增行
                </button>
              </div>
              <div className="flex gap-3">
                <button onClick={() => setView('upload')} className="px-5 py-2 text-gray-600 font-bold text-sm">取消</button>
                <button onClick={submitOrders} className="px-8 py-2 bg-blue-600 text-white rounded-xl font-bold text-sm shadow-lg shadow-blue-200 hover:bg-blue-700 active:scale-95 transition">
                  确认提交下单
                </button>
              </div>
            </div>

            <div className="bg-white rounded-2xl border shadow-sm overflow-hidden">
              <div className="overflow-x-auto max-h-[calc(100vh-250px)]">
                <table className="w-full text-sm border-collapse">
                  <thead className="sticky top-0 bg-gray-50 border-b z-10">
                    <tr>
                      <th className="p-4 text-center w-16 text-gray-400">#</th>
                      <th className="p-4 text-left w-24">操作</th>
                      {Object.values(FIELD_DEFINITIONS).map(d => (
                        <th key={d.label} className="p-4 text-left font-bold text-gray-500 whitespace-nowrap">
                          {d.label} {d.required && <span className="text-red-500">*</span>}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {orders.map((order, idx) => (
                      <tr key={order.id} className={`border-b hover:bg-gray-50/50 transition ${order.errors.length > 0 ? 'bg-red-50/30' : ''}`}>
                        <td className="p-4 text-center text-gray-400 font-mono">{idx + 1}</td>
                        <td className="p-4">
                          <button onClick={() => deleteRow(idx)} className="text-red-400 hover:text-red-600 p-1">
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </td>
                        {Object.keys(FIELD_DEFINITIONS).map(field => {
                          const error = order.errors.find((e: any) => e.field === field);
                          return (
                            <td key={field} className="p-2 relative group">
                              <input 
                                value={order[field] || ''}
                                onChange={(e) => updateCell(idx, field, e.target.value)}
                                className={`w-full min-w-[140px] p-2 rounded-lg border-2 transition focus:outline-none ${error ? 'border-red-400 bg-white ring-2 ring-red-100' : 'border-transparent focus:border-blue-400 focus:bg-white'}`}
                              />
                              {error && (
                                <div className="absolute hidden group-hover:block bottom-full left-2 mb-2 z-50 bg-gray-900 text-white text-[10px] py-1 px-2 rounded whitespace-nowrap pointer-events-none shadow-xl">
                                  {error.message}
                                </div>
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {view === 'history' && (
          <div className="max-w-5xl mx-auto space-y-6">
            <div className="bg-white p-6 rounded-2xl border shadow-sm flex items-center gap-4">
              <Search className="text-gray-400 w-5 h-5" />
              <input 
                placeholder="搜索运单信息..." 
                className="flex-1 bg-transparent border-none focus:ring-0 text-lg"
                value={searchTerm}
                onChange={e => setSearchTerm(e.target.value)}
              />
            </div>
            
            <div className="grid gap-4">
              {filteredHistory.map(batch => (
                <div key={batch.id} className="bg-white p-6 rounded-2xl border hover:shadow-md transition group">
                  <div className="flex justify-between items-center">
                    <div className="flex items-center gap-4">
                      <div className="bg-green-50 p-3 rounded-xl text-green-600">
                        <History className="w-6 h-6" />
                      </div>
                      <div>
                        <h4 className="font-bold text-lg">批量导入订单</h4>
                        <p className="text-sm text-gray-400">{batch.time} • {batch.count} 条记录</p>
                      </div>
                    </div>
                    <button className="flex items-center gap-2 text-sm font-bold text-blue-600 opacity-0 group-hover:opacity-100 transition px-4 py-2 bg-blue-50 rounded-lg">
                      查看详情 <Download className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              ))}
              {filteredHistory.length === 0 && (
                <div className="text-center py-20 text-gray-300">
                  <History className="w-16 h-16 mx-auto mb-4 opacity-10" />
                  <p className="font-bold">未找到相关历史记录</p>
                </div>
              )}
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
