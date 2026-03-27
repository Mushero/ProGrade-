import sqlite3
import sys
import threading
import uvicorn
import webview
import pandas as pd
import io
import os
import tkinter as tk
from tkinter import filedialog
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from datetime import datetime
from multiprocessing import freeze_support

# ==========================================
# 1. 数据库初始化 (SQLite)
# ==========================================
# DB_FILE = "data.db"
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

DB_FILE = os.path.join(application_path, "data.db")

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS classes (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS students (id INTEGER PRIMARY KEY AUTOINCREMENT, class_id INTEGER, student_id TEXT, name TEXT, score INTEGER DEFAULT 60)''')
    c.execute('''CREATE TABLE IF NOT EXISTS rules (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, value INTEGER, type TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS logs (id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER, rule_name TEXT, value INTEGER, created_at TEXT)''')
    
    c.execute("SELECT COUNT(*) FROM rules")
    if c.fetchone()[0] == 0:
        rules_data = [
            ('玩手机', -3, 'deduct'), ('课堂睡觉', -5, 'deduct'), ('交头接耳', -2, 'deduct'),
            ('主动答题', 2, 'add'), ('提出好问题', 3, 'add'), ('作业优秀', 2, 'add')
        ]
        c.executemany("INSERT INTO rules (name, value, type) VALUES (?, ?, ?)", rules_data)
    
    conn.commit()
    conn.close()

init_db()

# ==========================================
# 2. FastAPI 后端接口
# ==========================================
app = FastAPI()

class ScoreRequest(BaseModel):
    student_id: int
    rule_name: str
    value: int

class RuleRequest(BaseModel):
    name: str
    value: int
    type: str

@app.get("/api/data")
def get_data():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    classes = [dict(row) for row in c.execute("SELECT * FROM classes").fetchall()]
    students = [dict(row) for row in c.execute("SELECT * FROM students").fetchall()]
    rules = [dict(row) for row in c.execute("SELECT * FROM rules").fetchall()]
    conn.close()
    return {"classes": classes, "students": students, "rules": rules}

@app.delete("/api/classes/{class_id}")
def delete_class(class_id: int):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM logs WHERE student_id IN (SELECT id FROM students WHERE class_id = ?)", (class_id,))
    c.execute("DELETE FROM students WHERE class_id = ?", (class_id,))
    c.execute("DELETE FROM classes WHERE id = ?", (class_id,))
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.post("/api/score")
def update_score(req: ScoreRequest):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("UPDATE students SET score = score + ? WHERE id = ?", (req.value, req.student_id))
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("INSERT INTO logs (student_id, rule_name, value, created_at) VALUES (?, ?, ?, ?)", 
              (req.student_id, req.rule_name, req.value, now))
    conn.commit()
    c.execute("SELECT score FROM students WHERE id = ?", (req.student_id,))
    new_score = c.fetchone()[0]
    conn.close()
    return {"status": "success", "new_score": new_score}

@app.get("/api/logs/{student_id}")
def get_logs(student_id: int):
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    logs = [dict(row) for row in c.execute("SELECT * FROM logs WHERE student_id = ? ORDER BY id DESC", (student_id,)).fetchall()]
    conn.close()
    return logs

@app.post("/api/rules")
def add_rule(req: RuleRequest):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO rules (name, value, type) VALUES (?, ?, ?)", (req.name, req.value, req.type))
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.delete("/api/rules/{rule_id}")
def delete_rule(rule_id: int):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM rules WHERE id = ?", (rule_id,))
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.get("/api/export/{class_id}")
def export_excel(class_id: int):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT name FROM classes WHERE id = ?", (class_id,))
    row = c.fetchone()
    if not row:
        return {"status": "error", "message": "班级不存在"}
    class_name = row[0]
    
    query_students = f"SELECT student_id as 学号, name as 姓名, score as 平时总分 FROM students WHERE class_id = {class_id} ORDER BY student_id"
    df_students = pd.read_sql_query(query_students, conn)
    
    query_logs = f"""
        SELECT s.student_id as 学号, l.rule_name as 规则名称, l.value as 分值, l.created_at as 记录时间
        FROM logs l JOIN students s ON l.student_id = s.id
        WHERE s.class_id = {class_id} ORDER BY l.id ASC
    """
    df_logs = pd.read_sql_query(query_logs, conn)
    conn.close()

    if df_students.empty:
        return {"status": "error", "message": "该班级无学生数据，无法导出"}

    df_students.insert(0, '序号', range(1, len(df_students) + 1))

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=f"{class_name}_平时成绩归档.xlsx",
        title="选择保存位置",
        filetypes=[("Excel 表格", "*.xlsx")]
    )
    root.destroy()

    if not file_path:
        return {"status": "cancelled", "message": "已取消导出"}
        
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            safe_sheet_name = class_name[:31]
            df_students.to_excel(writer, sheet_name=safe_sheet_name, index=False, startrow=1)
            worksheet = writer.sheets[safe_sheet_name]
            worksheet.cell(row=1, column=1, value=class_name)
            
            if not df_logs.empty:
                log_sheet_name = (class_name[:26] + "_流水")
                df_logs.to_excel(writer, sheet_name=log_sheet_name, index=False)
                
        return {"status": "success", "message": f"成功导出至：{file_path}"}
    except Exception as e:
        return {"status": "error", "message": f"写入文件失败: {str(e)}"}

@app.post("/api/import")
async def import_excel(file: UploadFile = File(...)):
    contents = await file.read()
    try:
        df_dict = pd.read_excel(io.BytesIO(contents), sheet_name=None, header=None)
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        
        new_count, update_count, log_count = 0, 0, 0
        student_id_map = {} 
        
        for sheet_name, df in df_dict.items():
            if sheet_name.endswith("_流水") or len(df) < 2: continue
            
            class_name = str(df.iloc[0, 0]).strip()
            if pd.isna(class_name) or not class_name: class_name = sheet_name
            headers = [str(col).strip() for col in df.iloc[1].tolist()]
            data_df = df.iloc[2:].copy()
            data_df.columns = headers
            
            if not {'学号', '姓名'}.issubset(set(headers)): continue

            c.execute("SELECT id FROM classes WHERE name = ?", (class_name,))
            class_row = c.fetchone()
            class_id = class_row[0] if class_row else c.execute("INSERT INTO classes (name) VALUES (?)", (class_name,)).lastrowid
            
            for _, row in data_df.iterrows():
                if pd.isna(row.get('学号')): continue
                student_id = str(row['学号']).strip()
                student_name = str(row['姓名']).strip()
                score = 60
                if '平时总分' in headers and not pd.isna(row.get('平时总分')):
                    try: score = int(float(row['平时总分']))
                    except: pass
                
                c.execute("SELECT id FROM students WHERE student_id = ? AND class_id = ?", (student_id, class_id))
                student_row = c.fetchone()
                if student_row:
                    db_sid = student_row[0]
                    c.execute("UPDATE students SET score = ? WHERE id = ?", (score, db_sid))
                    update_count += 1
                else:
                    c.execute("INSERT INTO students (class_id, student_id, name, score) VALUES (?, ?, ?, ?)", 
                              (class_id, student_id, student_name, score))
                    db_sid = c.lastrowid
                    new_count += 1
                
                student_id_map[student_id] = db_sid

        for sheet_name, df in df_dict.items():
            if not sheet_name.endswith("_流水") or len(df) < 1: continue
            
            headers = [str(col).strip() for col in df.iloc[0].tolist()] 
            data_df = df.iloc[1:].copy()
            data_df.columns = headers
            
            if not {'学号', '规则名称', '分值', '记录时间'}.issubset(set(headers)): continue
            
            for _, row in data_df.iterrows():
                if pd.isna(row.get('学号')): continue
                student_id_str = str(row['学号']).strip()
                if student_id_str not in student_id_map: continue
                
                db_sid = student_id_map[student_id_str]
                rule_name = str(row['规则名称']).strip()
                created_at = str(row['记录时间']).strip()
                try: val = int(float(row['分值']))
                except: continue
                
                c.execute("SELECT id FROM logs WHERE student_id = ? AND rule_name = ? AND value = ? AND created_at = ?",
                          (db_sid, rule_name, val, created_at))
                if not c.fetchone():
                    c.execute("INSERT INTO logs (student_id, rule_name, value, created_at) VALUES (?, ?, ?, ?)",
                              (db_sid, rule_name, val, created_at))
                    log_count += 1

        conn.commit()
        conn.close()
        return {"status": "success", "message": f"处理成功！\n新增名单: {new_count} 人\n更新成绩: {update_count} 人\n恢复流水: {log_count} 条"}
    except Exception as e:
        return {"status": "error", "message": f"解析失败: {str(e)}"}

# ==========================================
# 3. 前端界面 (Vue3 + Tailwind)
# ==========================================
HTML_CONTENT = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>平时成绩管理系统 - Designed by suwnd</title>
    <script src="https://unpkg.com/vue@3/dist/vue.global.js"></script>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        .fade-enter-active, .fade-leave-active { transition: opacity 0.2s; }
        .fade-enter-from, .fade-leave-to { opacity: 0; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        /* 增加一个简单的脉冲动画给点名文字 */
        @keyframes fastpulse {
            0%, 100% { transform: scale(1); opacity: 1; }
            50% { transform: scale(1.05); opacity: 0.8; }
        }
        .animate-fastpulse { animation: fastpulse 0.2s infinite; }
    </style>
</head>
<body class="bg-gray-50 p-6 font-sans flex flex-col min-h-screen">
    <div id="app" class="max-w-7xl mx-auto w-full flex-grow">
        <div class="flex justify-between items-center mb-8 bg-white p-4 rounded-xl shadow-sm border-l-4 border-blue-600">
            <div class="flex space-x-3 items-center">
                <h1 class="text-2xl font-bold text-gray-800 mr-2 tracking-wide">课堂积分系统</h1>
                <select v-model="currentClass" class="p-2 border rounded-lg text-lg focus:ring-2 focus:ring-blue-500 min-w-[150px]">
                    <option v-if="classes.length === 0" value="null">请先导入名单</option>
                    <option v-for="c in classes" :key="c.id" :value="c.id">{{ c.name }}</option>
                </select>
                <button v-if="currentClass" @click="deleteCurrentClass" class="text-red-500 hover:bg-red-50 p-2 rounded-lg transition" title="学期结束删除本班级">
                    🗑️ 移除班级
                </button>
                <input v-model="searchQuery" type="text" placeholder="搜索学号/姓名..." class="p-2 border rounded-lg text-lg focus:ring-2 focus:ring-blue-500 ml-2">
            </div>
            
            <div class="flex items-center space-x-3">
                
                <button @click="startRandomCall" class="bg-purple-600 hover:bg-purple-700 text-white px-4 py-2 rounded-lg text-base transition font-semibold shadow-md flex items-center">
                    <span class="mr-1 text-xl">🎲</span> 随机点名
                </button>

                <button @click="showSettingsModal = true" class="bg-gray-800 hover:bg-gray-900 text-white px-4 py-2 rounded-lg text-base transition font-semibold shadow-md">
                    ⚙️ 规则设置
                </button>
                <input type="file" ref="fileInput" @change="handleFileUpload" accept=".xlsx, .xls" class="hidden">
                <button @click="$refs.fileInput.click()" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-base transition font-semibold shadow-md">
                    📥 导入表格
                </button>
                <button @click="exportData" class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-base transition font-semibold shadow-md">
                    📤 导出当前班级
                </button>
            </div>
        </div>

        <div v-if="filteredStudents.length === 0" class="text-center py-32 text-gray-400 text-xl flex flex-col items-center">
            <div class="text-6xl mb-4">📊</div>
            暂无当前班级数据，请点击右上角“导入表格”。
            <br><span class="text-sm mt-2 text-gray-400">（支持期末归档表格的一键无损导入）</span>
        </div>

        <div class="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-6">
            <div v-for="student in filteredStudents" :key="student.id" 
                 @click="openModal(student)"
                 class="bg-white rounded-xl shadow-sm border border-gray-100 p-6 cursor-pointer hover:shadow-lg hover:-translate-y-1 transition transform flex flex-col items-center justify-center relative overflow-hidden group">
                <div class="absolute inset-0 bg-blue-50 opacity-0 group-hover:opacity-100 transition-opacity"></div>
                <div class="relative z-10 text-gray-500 text-sm mb-1 font-mono">{{ student.student_id }}</div>
                <div class="relative z-10 text-2xl font-bold text-gray-800 mb-3">{{ student.name }}</div>
                <div class="relative z-10">
                    <span :class="{'text-green-600': student.score >= 60, 'text-red-600': student.score < 60}" class="text-4xl font-black">
                        {{ student.score }}
                    </span>
                </div>
            </div>
        </div>

        <transition name="fade">
            <div v-if="showRandomModal" class="fixed inset-0 bg-black bg-opacity-75 flex items-center justify-center z-[70]">
                <div class="bg-white rounded-[2rem] p-12 flex flex-col items-center justify-center shadow-2xl transform transition-all min-w-[500px] min-h-[350px]">
                    <h2 class="text-2xl font-bold text-gray-400 mb-6 tracking-widest">正在抽取天选之子...</h2>
                    <div :class="{'animate-fastpulse text-purple-600': isRolling, 'text-blue-600 scale-110': !isRolling}" 
                         class="text-8xl font-black tracking-widest transition-all duration-300">
                        {{ rollingName }}
                    </div>
                </div>
            </div>
        </transition>

        <transition name="fade">
            <div v-if="showScoreModal" class="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50" @click.self="showScoreModal = false">
                <div class="bg-white rounded-2xl flex max-w-5xl w-full shadow-2xl h-[600px] overflow-hidden">
                    <div class="w-2/3 p-8 border-r border-gray-100 flex flex-col">
                        <div class="flex justify-between items-center mb-6">
                            <h2 class="text-3xl font-bold text-gray-800">评分操作 - <span class="text-blue-600">{{ selectedStudent?.name }}</span></h2>
                        </div>
                        <div class="grid grid-cols-2 gap-6 overflow-y-auto pr-2">
                            <div>
                                <h3 class="text-xl font-bold text-red-600 mb-4">🔴 违纪扣分</h3>
                                <div class="space-y-3">
                                    <button v-for="rule in deductRules" :key="rule.id" @click="submitScore(rule)"
                                            class="w-full flex justify-between items-center bg-red-50 hover:bg-red-100 text-red-700 p-3 rounded-lg border border-red-200 transition">
                                        <span class="text-lg">{{ rule.name }}</span><span class="font-bold text-xl">{{ rule.value }}</span>
                                    </button>
                                </div>
                            </div>
                            <div>
                                <h3 class="text-xl font-bold text-green-600 mb-4">🟢 表现加分</h3>
                                <div class="space-y-3">
                                    <button v-for="rule in addRules" :key="rule.id" @click="submitScore(rule)"
                                            class="w-full flex justify-between items-center bg-green-50 hover:bg-green-100 text-green-700 p-3 rounded-lg border border-green-200 transition">
                                        <span class="text-lg">{{ rule.name }}</span><span class="font-bold text-xl">+{{ rule.value }}</span>
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="w-1/3 bg-gray-50 p-6 flex flex-col relative">
                        <button @click="showScoreModal = false" class="absolute top-4 right-6 text-gray-400 hover:text-red-500 text-3xl font-bold">&times;</button>
                        <h3 class="text-lg font-bold text-gray-700 mb-4 mt-2">📊 积分流水</h3>
                        <div class="flex-1 overflow-y-auto space-y-3 pr-2">
                            <div v-if="studentLogs.length === 0" class="text-gray-400 text-center mt-10 text-sm">暂无记录</div>
                            <div v-for="log in studentLogs" :key="log.id" class="bg-white p-3 rounded-lg shadow-sm border border-gray-100 text-sm">
                                <div class="flex justify-between items-center mb-1">
                                    <span class="font-bold text-gray-700">{{ log.rule_name }}</span>
                                    <span :class="log.value > 0 ? 'text-green-600' : 'text-red-600'" class="font-bold">
                                        {{ log.value > 0 ? '+'+log.value : log.value }}
                                    </span>
                                </div>
                                <div class="text-xs text-gray-400">{{ log.created_at }}</div>
                            </div>
                        </div>
                        <div class="mt-4 pt-4 border-t border-gray-200 flex justify-between items-center">
                            <span class="text-gray-500">当前总分</span><span class="text-3xl font-black text-blue-600">{{ selectedStudent?.score }}</span>
                        </div>
                    </div>
                </div>
            </div>
        </transition>

        <transition name="fade">
            <div v-if="showSettingsModal" class="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50" @click.self="showSettingsModal = false">
                <div class="bg-white rounded-2xl p-8 max-w-3xl w-full shadow-2xl h-[600px] flex flex-col">
                    <div class="flex justify-between items-center mb-6">
                        <h2 class="text-2xl font-bold text-gray-800">⚙️ 自定义计分项设置</h2>
                        <button @click="showSettingsModal = false" class="text-gray-500 hover:text-red-500 text-2xl font-bold">&times;</button>
                    </div>
                    <div class="grid grid-cols-2 gap-8 flex-1 overflow-y-auto">
                        <div class="bg-red-50/50 p-4 rounded-xl border border-red-100 flex flex-col">
                            <h3 class="text-lg font-bold text-red-700 mb-3">🔴 扣分项列表</h3>
                            <div class="space-y-2 flex-1 overflow-y-auto mb-4">
                                <div v-for="rule in deductRules" :key="rule.id" class="flex justify-between items-center bg-white p-2 rounded shadow-sm text-sm">
                                    <span>{{ rule.name }} ({{ rule.value }}分)</span>
                                    <button @click="deleteRule(rule.id)" class="text-red-400 hover:text-red-600 font-bold px-2">删除</button>
                                </div>
                            </div>
                            <div class="flex space-x-2 mt-auto">
                                <input v-model="newDeduct.name" placeholder="名称(如:迟到)" class="w-1/2 p-2 border rounded text-sm">
                                <input v-model.number="newDeduct.value" type="number" placeholder="分值(负数)" class="w-1/3 p-2 border rounded text-sm">
                                <button @click="addRule('deduct')" class="bg-red-500 hover:bg-red-600 text-white p-2 rounded text-sm w-1/4">添加</button>
                            </div>
                        </div>
                        <div class="bg-green-50/50 p-4 rounded-xl border border-green-100 flex flex-col">
                            <h3 class="text-lg font-bold text-green-700 mb-3">🟢 加分项列表</h3>
                            <div class="space-y-2 flex-1 overflow-y-auto mb-4">
                                <div v-for="rule in addRules" :key="rule.id" class="flex justify-between items-center bg-white p-2 rounded shadow-sm text-sm">
                                    <span>{{ rule.name }} (+{{ rule.value }}分)</span>
                                    <button @click="deleteRule(rule.id)" class="text-red-400 hover:text-red-600 font-bold px-2">删除</button>
                                </div>
                            </div>
                            <div class="flex space-x-2 mt-auto">
                                <input v-model="newAdd.name" placeholder="名称(如:全勤)" class="w-1/2 p-2 border rounded text-sm">
                                <input v-model.number="newAdd.value" type="number" placeholder="分值(正数)" class="w-1/3 p-2 border rounded text-sm">
                                <button @click="addRule('add')" class="bg-green-500 hover:bg-green-600 text-white p-2 rounded text-sm w-1/4">添加</button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </transition>
    </div>
    
    <footer class="mt-8 text-center text-gray-400 text-sm py-4 border-t border-gray-200">
        System Architecture Designed & Developed by <strong>suwnd</strong> © 2026. All Rights Reserved.
    </footer>

    <script>
        const { createApp, ref, computed, onMounted } = Vue;
        createApp({
            setup() {
                const classes = ref([]);
                const students = ref([]);
                const rules = ref([]);
                const currentClass = ref(null);
                const searchQuery = ref('');
                
                const showScoreModal = ref(false);
                const showSettingsModal = ref(false);
                const selectedStudent = ref(null);
                const studentLogs = ref([]);
                const fileInput = ref(null);

                const newDeduct = ref({ name: '', value: -1 });
                const newAdd = ref({ name: '', value: 1 });

                // 随机点名相关的响应式变量
                const showRandomModal = ref(false);
                const rollingName = ref('???');
                const isRolling = ref(false);
                let rollInterval = null;

                const loadData = async () => {
                    const res = await fetch('/api/data');
                    const data = await res.json();
                    classes.value = data.classes;
                    students.value = data.students;
                    rules.value = data.rules;
                    
                    if(!currentClass.value || !classes.value.find(c => c.id === currentClass.value)) {
                        currentClass.value = classes.value.length > 0 ? classes.value[0].id : null;
                    }
                };

                onMounted(loadData);

                const filteredStudents = computed(() => {
                    if (!students.value) return [];
                    return students.value.filter(s => s.class_id === currentClass.value && 
                        (s.name.includes(searchQuery.value) || s.student_id.includes(searchQuery.value)));
                });

                const deductRules = computed(() => rules.value.filter(r => r.type === 'deduct'));
                const addRules = computed(() => rules.value.filter(r => r.type === 'add'));

                const fetchLogs = async (studentId) => {
                    const res = await fetch(`/api/logs/${studentId}`);
                    studentLogs.value = await res.json();
                };

                const openModal = async (student) => {
                    selectedStudent.value = student;
                    showScoreModal.value = true;
                    await fetchLogs(student.id);
                };

                // 🌟 新增：随机点名核心逻辑
                const startRandomCall = () => {
                    const list = filteredStudents.value;
                    if (list.length === 0) {
                        alert("当前列表没有学生，请先检查班级数据！");
                        return;
                    }

                    showRandomModal.value = true;
                    isRolling.value = true;
                    
                    // 名字快速滚动动画 (每50ms切换一次)
                    rollInterval = setInterval(() => {
                        const randomIndex = Math.floor(Math.random() * list.length);
                        rollingName.value = list[randomIndex].name;
                    }, 50);

                    // 2秒后停止滚动并锁定最终学生
                    setTimeout(() => {
                        clearInterval(rollInterval);
                        isRolling.value = false;
                        
                        const finalIndex = Math.floor(Math.random() * list.length);
                        const winner = list[finalIndex];
                        rollingName.value = winner.name;
                        
                        // 停顿 1.2 秒让大家看清名字，然后自动打开该学生的打分界面
                        setTimeout(() => {
                            showRandomModal.value = false;
                            openModal(winner);
                        }, 1200);

                    }, 2000);
                };

                const submitScore = async (rule) => {
                    const res = await fetch('/api/score', {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({ student_id: selectedStudent.value.id, rule_name: rule.name, value: rule.value })
                    });
                    const data = await res.json();
                    if(data.status === 'success') {
                        selectedStudent.value.score = data.new_score;
                        await fetchLogs(selectedStudent.value.id);
                    }
                };

                const addRule = async (type) => {
                    const payload = type === 'deduct' ? newDeduct.value : newAdd.value;
                    if(!payload.name || !payload.value) return alert("名称和分值不能为空");
                    if(type === 'deduct' && payload.value > 0) payload.value = -payload.value;
                    if(type === 'add' && payload.value < 0) payload.value = Math.abs(payload.value);

                    await fetch('/api/rules', {
                        method: 'POST',
                        headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({ name: payload.name, value: payload.value, type: type })
                    });
                    
                    if(type === 'deduct') newDeduct.value = {name: '', value: -1};
                    else newAdd.value = {name: '', value: 1};
                    await loadData();
                };

                const deleteRule = async (ruleId) => {
                    if(!confirm("确定要删除该计分项吗？（不影响已有学生的历史记录）")) return;
                    await fetch(`/api/rules/${ruleId}`, { method: 'DELETE' });
                    await loadData();
                };
                
                const deleteCurrentClass = async () => {
                    const classObj = classes.value.find(c => c.id === currentClass.value);
                    if(!classObj) return;
                    if(!confirm(`⚠️ 警告：确定要删除【${classObj.name}】吗？\n这会彻底清空该班级所有学生和流水记录！\n请确保你已经导出备份！`)) return;
                    
                    await fetch(`/api/classes/${currentClass.value}`, { method: 'DELETE' });
                    await loadData();
                };

                const handleFileUpload = async (event) => {
                    const file = event.target.files[0];
                    if (!file) return;
                    const formData = new FormData();
                    formData.append('file', file);
                    try {
                        const res = await fetch('/api/import', { method: 'POST', body: formData });
                        const data = await res.json();
                        alert(data.message);
                        if (data.status === 'success') await loadData();
                    } catch (error) {
                        alert('上传请求发生错误');
                    }
                    event.target.value = '';
                };

                const exportData = async () => {
                    if (!currentClass.value) return alert("请先选择要导出的班级");
                    try {
                        const res = await fetch(`/api/export/${currentClass.value}`);
                        const data = await res.json();
                        if(data.status !== 'cancelled') {
                            alert(data.message);
                        }
                    } catch(e) {
                        alert('请求导出失败');
                    }
                };

                return { 
                    classes, currentClass, searchQuery, filteredStudents, 
                    showScoreModal, showSettingsModal, selectedStudent, studentLogs,
                    deductRules, addRules, fileInput, newDeduct, newAdd,
                    showRandomModal, rollingName, isRolling, startRandomCall, /* 对外暴露点名函数和变量 */
                    openModal, submitScore, addRule, deleteRule, deleteCurrentClass, handleFileUpload, exportData 
                }
            }
        }).mount('#app');
    </script>
</body>
</html>
"""

@app.get("/")
def read_root():
    return HTMLResponse(content=HTML_CONTENT)

def run_api():
    # uvicorn.run(app, host="127.0.0.1", port=8000, log_level="error")
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="error", reload=False, workers=1)

if __name__ == '__main__':
    freeze_support()

    api_thread = threading.Thread(target=run_api, daemon=True)
    api_thread.start()
    webview.create_window('平时成绩管理系统', 'http://127.0.0.1:8000', width=1280, height=800, min_size=(1000, 700))
    webview.start()