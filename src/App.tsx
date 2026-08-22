import {
  ArrowUpRight,
  BarChart3,
  BookOpen,
  CalendarDays,
  Check,
  ChevronDown,
  Clock3,
  Download,
  Flame,
  GraduationCap,
  LayoutDashboard,
  ListChecks,
  Plus,
  Target,
  TrendingUp,
  Users,
  X,
} from "lucide-react";
import type { FormEvent } from "react";
import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import "./App.css";

type ExcelFileHandle = FileSystemFileHandle;

type Entry = {
  id: number;
  student: string;
  date: string;
  learned: string;
  practiced: string;
  questions: number;
  remarks: string;
  cuetPaperYear: string;
  cuetPaperScore: string;
  clatPaperYear: string;
  clatPaperScore: string;
  duration: number;
  confidence: string;
  nextFocus: string;
};

const students = [
  { name: "Bhavisya ", initials: "AM", color: "coral", role: "CUET + CLAT" },
  { name: "Himanshu", initials: "DS", color: "mint", role: "CUET focus" },
  { name: "Krish", initials: "KR", color: "yellow", role: "CLAT focus" },
];

const today = new Date().toISOString().slice(0, 10);
const DEFAULT_ONLINE_SHEET_URL =
  "https://script.google.com/macros/s/AKfycbxtbweyyXGpnrxkL9BG7IPaxjlwe1W8r_8BxPSP3hji915JgNPmjyfi63RGlqMwYG-paQ/exec";
function App() {
  const [entries, setEntries] = useState<Entry[]>(() => {
    const saved = localStorage.getItem("prep-log-entries");
    if (!saved) return [];
    return JSON.parse(saved)
      .map(
        (
          entry: Entry & {
            exam?: "CUET" | "CLAT";
            paperYear?: string;
            paperScore?: string;
            pyqYear?: string;
            pyqScore?: string;
          },
        ) => ({
          ...entry,
          cuetPaperYear:
            entry.cuetPaperYear ??
            (entry.exam === "CUET"
              ? (entry.paperYear ?? entry.pyqYear ?? "—")
              : ""),
          cuetPaperScore:
            entry.cuetPaperScore ??
            (entry.exam === "CUET"
              ? (entry.paperScore ?? entry.pyqScore ?? "")
              : ""),
          clatPaperYear:
            entry.clatPaperYear ??
            (entry.exam === "CLAT"
              ? (entry.paperYear ?? entry.pyqYear ?? "—")
              : ""),
          clatPaperScore:
            entry.clatPaperScore ??
            (entry.exam === "CLAT"
              ? (entry.paperScore ?? entry.pyqScore ?? "")
              : ""),
        }),
      )
      .filter(
        (entry: Entry) =>
          ![1, 2, 3, 4].includes(Number(entry.id)) &&
          String(entry.id) !== "connection-test-20260822",
      );
  });
  const [activeStudent, setActiveStudent] = useState("All students");
  const [isFormOpen, setIsFormOpen] = useState(false);
  const [savedNotice, setSavedNotice] = useState(false);
  const [excelHandle, setExcelHandle] = useState<ExcelFileHandle | null>(null);
  const [excelFileName, setExcelFileName] = useState("");
  const [isSpreadsheetOpen, setIsSpreadsheetOpen] = useState(false);
  const [onlineSheetUrl, setOnlineSheetUrl] = useState(
    () =>
      DEFAULT_ONLINE_SHEET_URL ||
      localStorage.getItem("prep-log-sheet-url") ||
      "",
  );
  const [onlineSyncing, setOnlineSyncing] = useState(false);
  const [form, setForm] = useState({
    student: students[0].name,
    date: today,
    learned: "",
    practiced: "",
    questions: "",
    remarks: "",
    cuetPaperYear: "",
    cuetPaperScore: "",
    clatPaperYear: "",
    clatPaperScore: "",
    duration: "120",
    confidence: "Steady",
    nextFocus: "",
  });

  const filteredEntries = useMemo(
    () =>
      entries.filter(
        (entry) =>
          activeStudent === "All students" || entry.student === activeStudent,
      ),
    [entries, activeStudent],
  );
  const totalQuestions = filteredEntries.reduce(
    (sum, entry) => sum + entry.questions,
    0,
  );
  const totalHours = Math.round(
    filteredEntries.reduce((sum, entry) => sum + entry.duration, 0) / 60,
  );
  const paperScores = filteredEntries
    .flatMap((entry) => [entry.cuetPaperScore, entry.clatPaperScore])
    .filter((score) => score && score.includes("/"));
  const averageScore = paperScores.length
    ? Math.round(
        paperScores.reduce(
          (sum, score) =>
            sum +
            (Number(score.split("/")[0]) / Number(score.split("/")[1])) * 100,
          0,
        ) / paperScores.length,
      )
    : 0;

  useEffect(() => {
    if (onlineSheetUrl) void loadOnlineSheet(onlineSheetUrl);
  }, [onlineSheetUrl]);

  function updateForm(field: string, value: string) {
    setForm((current) => ({ ...current, [field]: value }));
  }

  function saveEntry(event: FormEvent) {
    event.preventDefault();
    const formatScore = (score: string, total: string) =>
      score && !score.includes("/") ? `${score} / ${total}` : score;
    const newEntry: Entry = {
      ...form,
      id: Date.now(),
      questions: Number(form.questions) || 0,
      duration: Number(form.duration) || 0,
      cuetPaperYear: form.cuetPaperYear,
      cuetPaperScore: formatScore(form.cuetPaperScore, "250"),
      clatPaperYear: form.clatPaperYear,
      clatPaperScore: formatScore(form.clatPaperScore, "120"),
    };
    const nextEntries = [newEntry, ...entries];
    setEntries(nextEntries);
    localStorage.setItem("prep-log-entries", JSON.stringify(nextEntries));
    void syncExcelFile(nextEntries);
    void syncOnlineSheet(newEntry);
    setSavedNotice(true);
    setIsFormOpen(false);
    setForm({
      student: students[0].name,
      date: today,
      learned: "",
      practiced: "",
      questions: "",
      remarks: "",
      cuetPaperYear: "",
      cuetPaperScore: "",
      clatPaperYear: "",
      clatPaperScore: "",
      duration: "120",
      confidence: "Steady",
      nextFocus: "",
    });
    window.setTimeout(() => setSavedNotice(false), 2600);
  }

  async function connectExcelFile() {
    const picker = (
      window as Window & {
        showSaveFilePicker?: (options?: object) => Promise<ExcelFileHandle>;
      }
    ).showSaveFilePicker;
    if (!picker) {
      setSavedNotice(true);
      window.setTimeout(() => setSavedNotice(false), 3000);
      return;
    }
    const handle = await picker({
      suggestedName: "cuet-clat-study-log.xlsx",
      types: [
        {
          description: "Excel workbook",
          accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
              [".xlsx"],
          },
        },
      ],
    });
    setExcelHandle(handle);
    setExcelFileName(handle.name);
    await writeExcelFile(handle, entries);
    setSavedNotice(true);
    window.setTimeout(() => setSavedNotice(false), 2600);
  }

  async function writeExcelFile(handle: ExcelFileHandle, data: Entry[]) {
    const rows = data.map((entry) => ({
      Student: entry.student,
      Date: entry.date,
      "New topics learned": entry.learned,
      "Topics practiced": entry.practiced,
      Questions: entry.questions,
      Remarks: entry.remarks,
      "CUET paper year": entry.cuetPaperYear,
      "CUET paper score": entry.cuetPaperScore,
      "CLAT paper year": entry.clatPaperYear,
      "CLAT paper score": entry.clatPaperScore,
      "Study minutes": entry.duration,
      Confidence: entry.confidence,
      "Next focus": entry.nextFocus,
    }));
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Daily log");
    const output = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const writable = await handle.createWritable();
    await writable.write(output);
    await writable.close();
  }

  async function syncExcelFile(data: Entry[]) {
    if (!excelHandle) return;
    try {
      await writeExcelFile(excelHandle, data);
    } catch {
      setExcelHandle(null);
      setExcelFileName("");
    }
  }

  async function connectOnlineSheet() {
    const url = window
      .prompt(
        "Paste the deployed Google Apps Script web app URL:",
        onlineSheetUrl,
      )
      ?.trim();
    if (!url) return;
    setOnlineSheetUrl(url);
    localStorage.setItem("prep-log-sheet-url", url);
    setOnlineSyncing(true);
    try {
      await loadOnlineSheet(url);
      setSavedNotice(true);
      window.setTimeout(() => setSavedNotice(false), 2600);
    } finally {
      setOnlineSyncing(false);
    }
  }

  async function loadOnlineSheet(url: string) {
    const response = await fetch(url);
    const remote = (await response.json()) as { entries?: Entry[] };
    if (remote.entries?.length) {
      setEntries(remote.entries);
      localStorage.setItem("prep-log-entries", JSON.stringify(remote.entries));
    }
  }

  async function syncOnlineSheet(entry: Entry) {
    if (!onlineSheetUrl) return;
    setOnlineSyncing(true);
    try {
      await fetch(onlineSheetUrl, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ entry }),
      });
      setSavedNotice(true);
      window.setTimeout(() => setSavedNotice(false), 2600);
    } catch {
      setSavedNotice(true);
      window.setTimeout(() => setSavedNotice(false), 4000);
    } finally {
      setOnlineSyncing(false);
    }
  }

  function exportCsv() {
    const headers = [
      "Student",
      "Date",
      "New topics learned",
      "Topics practiced",
      "Questions",
      "Remarks",
      "CUET paper year",
      "CUET paper score",
      "CLAT paper year",
      "CLAT paper score",
      "Study minutes",
      "Confidence",
      "Next focus",
    ];
    const rows = entries.map((entry) => [
      entry.student,
      entry.date,
      entry.learned,
      entry.practiced,
      entry.questions,
      entry.remarks,
      entry.cuetPaperYear,
      entry.cuetPaperScore,
      entry.clatPaperYear,
      entry.clatPaperScore,
      entry.duration,
      entry.confidence,
      entry.nextFocus,
    ]);
    const csv = [headers, ...rows]
      .map((row) =>
        row.map((cell) => `"${String(cell).replaceAll('"', '""')}"`).join(","),
      )
      .join("\n");
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
    link.download = "cuet-clat-prep-log.csv";
    link.click();
    URL.revokeObjectURL(link.href);
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <span className="brand-mark">
            <GraduationCap size={20} />
          </span>
          <span>
            study<span className="brand-dot">.</span>log
          </span>
        </div>
        <div className="sidebar-label">Workspace</div>
        <nav>
          <button className="nav-item active">
            <LayoutDashboard size={17} /> Overview
          </button>
          <button className="nav-item">
            <ListChecks size={17} /> Daily entries{" "}
            <span className="nav-count">{entries.length}</span>
          </button>
          <button className="nav-item">
            <BarChart3 size={17} /> Insights
          </button>
        </nav>
        <div className="sidebar-label student-label">
          Students{" "}
          <button className="icon-button" aria-label="Add student">
            <Plus size={15} />
          </button>
        </div>
        <div className="student-list">
          {students.map((student) => (
            <button
              className="student-nav"
              key={student.name}
              onClick={() => setActiveStudent(student.name)}
            >
              <span className={`avatar ${student.color}`}>
                {student.initials}
              </span>
              <span>
                <strong>{student.name}</strong>
                <small>{student.role}</small>
              </span>
              <span className="status-dot" />
            </button>
          ))}
        </div>
        <div className="sidebar-bottom">
          <div className="tip-card">
            <Target size={18} />
            <div>
              <strong>Small steps, daily.</strong>
              <button
                className="ghost-button"
                onClick={() => void connectOnlineSheet()}
              >
                <TrendingUp size={16} />{" "}
                {onlineSyncing
                  ? "Syncing..."
                  : onlineSheetUrl
                    ? "Online sheet connected"
                    : "Connect online sheet"}
              </button>
              <span>Consistency compounds.</span>
            </div>
          </div>
          <div className="profile">
            <span className="avatar profile-avatar">PS</span>
            <span>
              <strong>Parent / mentor</strong>
              <small>Workspace admin</small>
            </span>
            <ChevronDown size={15} />
          </div>
        </div>
      </aside>

      <main className="main-content">
        <header className="topbar">
          <div className="crumb">
            <span>Workspace</span>
            <span>/</span>
            <strong>Overview</strong>
          </div>
          <div className="top-actions">
            <button
              className="ghost-button"
              onClick={() => void connectExcelFile()}
            >
              <Download size={16} />{" "}
              {excelFileName ? "Excel connected" : "Connect Excel"}
            </button>
            <button className="ghost-button" onClick={exportCsv}>
              <Download size={16} /> Download CSV
            </button>
            <button
              className="primary-button"
              onClick={() => setIsFormOpen(true)}
            >
              <Plus size={17} /> Log today
            </button>
          </div>
        </header>
        <section className="page-heading">
          <div>
            <p className="eyebrow">Saturday, 22 August 2026</p>
            <h1>
              Good morning, <em>mentor.</em>
            </h1>
            <p className="heading-copy">
              A clear view of the work that moves your students forward.
            </p>
          </div>
          <div className="streak">
            <Flame size={22} fill="currentColor" />
            <span>
              <strong>8 day</strong>
              <small>team streak</small>
            </span>
          </div>
        </section>

        <section className="filter-row">
          <div className="segmented">
            {["All students", ...students.map((s) => s.name)].map((student) => (
              <button
                key={student}
                className={activeStudent === student ? "selected" : ""}
                onClick={() => setActiveStudent(student)}
              >
                {student === "All students" ? <Users size={15} /> : null}
                {student}
              </button>
            ))}
          </div>
        </section>

        <section className="stat-grid">
          <div className="stat-card">
            <div className="stat-icon coral-icon">
              <ListChecks size={19} />
            </div>
            <span>Entries logged</span>
            <strong>
              {filteredEntries.length}
              <small> this view</small>
            </strong>
            <div className="stat-foot up">
              <TrendingUp size={14} /> +3 this week
            </div>
          </div>
          <div className="stat-card">
            <div className="stat-icon blue-icon">
              <BookOpen size={19} />
            </div>
            <span>Questions practised</span>
            <strong>
              {totalQuestions}
              <small> total</small>
            </strong>
            <div className="stat-foot up">
              <TrendingUp size={14} /> 12% vs last week
            </div>
          </div>
          <div className="stat-card">
            <div className="stat-icon yellow-icon">
              <Clock3 size={19} />
            </div>
            <span>Study time</span>
            <strong>
              {totalHours}
              <small> hrs</small>
            </strong>
            <div className="stat-foot">Across all sessions</div>
          </div>
          <div className="stat-card">
            <div className="stat-icon mint-icon">
              <Target size={19} />
            </div>
            <span>Avg. paper score</span>
            <strong>{averageScore}%</strong>
            <div className="stat-foot up">
              <TrendingUp size={14} /> Keep pushing
            </div>
          </div>
        </section>

        <div className="content-grid">
          <section className="activity-panel panel">
            <div className="panel-header">
              <div>
                <h2>Recent activity</h2>
                <p>The latest notes from each study desk.</p>
              </div>
              <button className="text-button">
                View all <ArrowUpRight size={15} />
              </button>
            </div>
            <div className="activity-list">
              {filteredEntries.slice(0, 4).map((entry) => (
                <article className="activity-item" key={entry.id}>
                  <span
                    className={`avatar ${students.find((s) => s.name === entry.student)?.color}`}
                  >
                    {students.find((s) => s.name === entry.student)?.initials}
                  </span>
                  <div className="activity-main">
                    <div className="activity-top">
                      <strong>{entry.student}</strong>
                      {entry.cuetPaperScore && (
                        <span className="exam-tag cuet">CUET paper</span>
                      )}
                      {entry.clatPaperScore && (
                        <span className="exam-tag clat">CLAT paper</span>
                      )}
                      <time>
                        {entry.date === "2026-08-21" ? "Yesterday" : entry.date}
                      </time>
                    </div>
                    <p>
                      <b>{entry.learned}</b> <span>·</span> {entry.questions}{" "}
                      questions practised
                    </p>
                    <small>
                      <span className="mini-check">
                        <Check size={11} />
                      </span>{" "}
                      {entry.practiced}
                    </small>
                  </div>
                  <div className="score">
                    <strong>
                      {entry.cuetPaperScore || entry.clatPaperScore
                        ? "View"
                        : "—"}
                    </strong>
                    <small>paper details</small>
                  </div>
                </article>
              ))}
            </div>
          </section>
          <section className="focus-panel panel">
            <div className="panel-header">
              <div>
                <h2>Next focus</h2>
                <p>What to pick up tomorrow.</p>
              </div>
              <Target size={19} className="panel-icon" />
            </div>
            <div className="focus-list">
              {students.map((student, index) => {
                const entry = entries.find(
                  (item) => item.student === student.name,
                );
                return (
                  <div className="focus-item" key={student.name}>
                    <span className={`avatar ${student.color}`}>
                      {student.initials}
                    </span>
                    <div>
                      <strong>{student.name}</strong>
                      <p>
                        {entry?.nextFocus ||
                          [
                            "Reading comprehension",
                            "Algebra revision",
                            "Legal reasoning",
                          ][index]}
                      </p>
                    </div>
                    <span className="focus-arrow">→</span>
                  </div>
                );
              })}
            </div>
            <div className="insight">
              <span className="insight-spark">✦</span>
              <p>
                <strong>One useful signal</strong>{" "}
                {activeStudent === "All students"
                  ? "Your team is strongest on consistency this week."
                  : `${activeStudent} has logged ${filteredEntries.length} recent sessions.`}
              </p>
            </div>
          </section>
        </div>

        <section className="log-banner">
          <div className="banner-icon">
            <CalendarDays size={22} />
          </div>
          <div>
            <strong>Keep the picture complete.</strong>
            <span>
              One short entry per student makes weekly patterns visible.
            </span>
          </div>
          <button
            className="outline-button"
            onClick={() => setIsSpreadsheetOpen((open) => !open)}
          >
            {isSpreadsheetOpen ? "Hide spreadsheet" : "View spreadsheet"}{" "}
            <ArrowUpRight size={16} />
          </button>
        </section>
        {isSpreadsheetOpen && (
          <section className="spreadsheet-panel panel">
            <div className="panel-header">
              <div>
                <h2>Daily log spreadsheet</h2>
                <p>
                  {excelFileName
                    ? `Auto-syncing to ${excelFileName}`
                    : "Connect an Excel file to keep this sheet synced automatically."}
                </p>
              </div>
              <button
                className="text-button"
                onClick={() => void connectExcelFile()}
              >
                {excelFileName ? "Change file" : "Connect file"}{" "}
                <Download size={15} />
              </button>
            </div>
            <div className="table-scroll">
              <table>
                <thead>
                  <tr>
                    <th>Date</th>
                    <th>Student</th>
                    <th>New topics learned</th>
                    <th>Practised</th>
                    <th>Questions</th>
                    <th>CUET paper</th>
                    <th>CLAT paper</th>
                    <th>Minutes</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredEntries.map((entry) => (
                    <tr key={entry.id}>
                      <td>{entry.date}</td>
                      <td>
                        <strong>{entry.student}</strong>
                      </td>
                      <td>{entry.learned}</td>
                      <td>{entry.practiced}</td>
                      <td>{entry.questions}</td>
                      <td>
                        {entry.cuetPaperYear
                          ? `${entry.cuetPaperYear} · ${entry.cuetPaperScore || "—"}`
                          : "—"}
                      </td>
                      <td>
                        {entry.clatPaperYear
                          ? `${entry.clatPaperYear} · ${entry.clatPaperScore || "—"}`
                          : "—"}
                      </td>
                      <td>{entry.duration}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        )}
      </main>

      {isFormOpen && (
        <div
          className="modal-backdrop"
          onMouseDown={(event) =>
            event.target === event.currentTarget && setIsFormOpen(false)
          }
        >
          <form className="entry-modal" onSubmit={saveEntry}>
            <div className="modal-header">
              <div>
                <p className="eyebrow">New daily summary</p>
                <h2>Log a study day</h2>
              </div>
              <button
                type="button"
                className="close-button"
                onClick={() => setIsFormOpen(false)}
                aria-label="Close form"
              >
                <X size={19} />
              </button>
            </div>
            <div className="form-grid">
              <div className="form-section wide">
                <strong>Daily study log</strong>
                <span>Capture the work done today, regardless of exam.</span>
              </div>
              <label>
                Student
                <select
                  value={form.student}
                  onChange={(e) => updateForm("student", e.target.value)}
                >
                  {students.map((s) => (
                    <option key={s.name}>{s.name}</option>
                  ))}
                </select>
              </label>
              <label>
                Date
                <input
                  type="date"
                  value={form.date}
                  onChange={(e) => updateForm("date", e.target.value)}
                  required
                />
              </label>
              <label>
                Study time (minutes)
                <input
                  type="number"
                  min="0"
                  value={form.duration}
                  onChange={(e) => updateForm("duration", e.target.value)}
                />
              </label>
              <label className="wide">
                New topics learned
                <input
                  value={form.learned}
                  onChange={(e) => updateForm("learned", e.target.value)}
                  placeholder="e.g. Constitutional law: Fundamental Rights"
                  required
                />
              </label>
              <label className="wide">
                Topics practised{" "}
                <span className="label-hint">separate with commas</span>
                <input
                  value={form.practiced}
                  onChange={(e) => updateForm("practiced", e.target.value)}
                  placeholder="e.g. Legal reasoning, English comprehension"
                  required
                />
              </label>
              <label>
                Questions practised
                <input
                  type="number"
                  min="0"
                  value={form.questions}
                  onChange={(e) => updateForm("questions", e.target.value)}
                  placeholder="0"
                  required
                />
              </label>
              <label>
                Confidence
                <select
                  value={form.confidence}
                  onChange={(e) => updateForm("confidence", e.target.value)}
                >
                  <option>Strong</option>
                  <option>Steady</option>
                  <option>Needs review</option>
                </select>
              </label>
              <div className="form-section wide paper-section">
                <strong>
                  Papers solved <span>(optional)</span>
                </strong>
                <span>
                  Add details for both papers when they were practised today.
                </span>
              </div>
              <label>
                CUET paper year
                <input
                  value={form.cuetPaperYear}
                  onChange={(e) => updateForm("cuetPaperYear", e.target.value)}
                  placeholder="2024"
                />
              </label>
              <label>
                CUET paper score{" "}
                <span className="label-hint">e.g. 164 / 250</span>
                <input
                  value={form.cuetPaperScore}
                  onChange={(e) => updateForm("cuetPaperScore", e.target.value)}
                  placeholder="Score / 250"
                />
              </label>
              <label>
                CLAT paper year
                <input
                  value={form.clatPaperYear}
                  onChange={(e) => updateForm("clatPaperYear", e.target.value)}
                  placeholder="2024"
                />
              </label>
              <label>
                CLAT paper score{" "}
                <span className="label-hint">e.g. 78 / 120</span>
                <input
                  value={form.clatPaperScore}
                  onChange={(e) => updateForm("clatPaperScore", e.target.value)}
                  placeholder="Score / 120"
                />
              </label>
              <label className="wide">
                Next focus
                <input
                  value={form.nextFocus}
                  onChange={(e) => updateForm("nextFocus", e.target.value)}
                  placeholder="What should be picked up next?"
                />
              </label>
              <label className="wide">
                Extra remarks
                <textarea
                  rows={3}
                  value={form.remarks}
                  onChange={(e) => updateForm("remarks", e.target.value)}
                  placeholder="A quick note on what went well or needs attention..."
                />
              </label>
            </div>
            <div className="modal-actions">
              <button
                type="button"
                className="ghost-button"
                onClick={() => setIsFormOpen(false)}
              >
                Cancel
              </button>
              <button type="submit" className="primary-button">
                <Check size={16} /> Save summary
              </button>
            </div>
          </form>
        </div>
      )}
      {savedNotice && (
        <div className="toast">
          <Check size={17} /> Daily summary saved
        </div>
      )}
    </div>
  );
}

export default App;
