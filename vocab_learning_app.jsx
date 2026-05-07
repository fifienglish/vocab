import React, { useMemo, useState } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { Volume2, Check, X, ArrowRight, RotateCcw, Eye, EyeOff, Upload, BookOpen, AlertCircle, ListChecks } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import * as XLSX from "xlsx";

const SAMPLE_WORDS = [
  {
    word: "booklet",
    phonetic: "/ˈbʊklət/",
    partOfSpeech: "n.",
    meaning: "小册子",
    shortMeaning: "小册子",
    example: "The tourist information center gave us a free ____ about local attractions.",
    options: ["手册", "小册子", "传单", "指南"],
    correctOption: "小册子",
    explanation: "booklet 是小册子，通常用于宣传、介绍或说明。",
    hint: "book + let：像一本小小的 book。",
  },
  {
    word: "reluctant",
    phonetic: "/rɪˈlʌktənt/",
    partOfSpeech: "adj.",
    meaning: "不情愿的；勉强的",
    shortMeaning: "不情愿的",
    example: "She was ____ to speak in front of the class because she felt nervous.",
    options: ["兴奋的", "不情愿的", "诚实的", "困惑的"],
    correctOption: "不情愿的",
    explanation: "reluctant 表示因为不想做或有顾虑，所以表现出勉强、不情愿。",
    hint: "可以联想：really not want to do something。",
  },
  {
    word: "fragile",
    phonetic: "/ˈfrædʒaɪl/",
    partOfSpeech: "adj.",
    meaning: "易碎的；脆弱的",
    shortMeaning: "易碎的",
    example: "Please be careful with the glass vase because it is very ____.",
    options: ["易碎的", "昂贵的", "普通的", "沉重的"],
    correctOption: "易碎的",
    explanation: "fragile 常用于描述玻璃、陶瓷等容易破碎的东西，也可形容人或关系很脆弱。",
    hint: "frag- 有 broken 的感觉，可联想到 fracture（骨折）。",
  },
];

const STAGES = [
  { key: "context", label: "L1 · 语境理解", color: "blue" },
  { key: "detail", label: "单词详情", color: "blue" },
  { key: "sound", label: "拼写 · 听音写词", color: "green" },
  { key: "meaning", label: "拼写 · 看义写词", color: "green" },
];

function normaliseWord(row) {
  const word = String(row.word || "").trim();
  const meaning = String(row.meaning || row.shortMeaning || "").trim();
  if (!word || !meaning) return null;

  const options = Array.isArray(row.options)
    ? row.options
    : String(row.options || "")
        .split(/[|,，、]/)
        .map((item) => item.trim())
        .filter(Boolean);

  return {
    word,
    phonetic: String(row.phonetic || "").trim() || "/ 点击音频听发音 /",
    partOfSpeech: String(row.partOfSpeech || row.pos || "").trim(),
    meaning,
    shortMeaning: String(row.shortMeaning || meaning).trim(),
    example: String(row.example || `I learned the word ____ today.`).trim(),
    options: options.length >= 4 ? options.slice(0, 4) : Array.from(new Set([meaning, "相反含义", "近似干扰项", "无关含义"])).slice(0, 4),
    correctOption: String(row.correctOption || row.shortMeaning || meaning).trim(),
    explanation: String(row.explanation || `${word} 的意思是：${meaning}`).trim(),
    hint: String(row.hint || "可以结合例句、发音和词根词缀进行记忆。 ").trim(),
  };
}

function parseCSV(text) {
  const lines = text.split(/?
/).filter((line) => line.trim());
  if (lines.length < 2) return [];
  const splitLine = (line) => line.split(",").map((cell) => cell.trim().replace(/^"|"$/g, ""));
  const headers = splitLine(lines[0]);
  return lines
    .slice(1)
    .map((line) => {
      const cells = splitLine(line);
      return headers.reduce((obj, header, index) => ({ ...obj, [header]: cells[index] || "" }), {});
    })
    .map(normaliseWord)
    .filter(Boolean);
}

function parseExcel(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
  return rows.map(normaliseWord).filter(Boolean);
}

function speak(text) {
  if (typeof window === "undefined" || !window.speechSynthesis) return;
  const utterance = new SpeechSynthesisUtterance(text);
  utterance.lang = "en-US";
  utterance.rate = 0.82;
  window.speechSynthesis.cancel();
  window.speechSynthesis.speak(utterance);
}

function Pill({ children, tone = "blue" }) {
  const styles = tone === "green" ? "bg-green-100 text-green-700" : tone === "red" ? "bg-red-100 text-red-700" : "bg-blue-100 text-blue-700";
  return <span className={`inline-flex rounded-full px-5 py-2 text-sm font-semibold ${styles}`}>{children}</span>;
}

function AudioButton({ word }) {
  return (
    <button
      onClick={() => speak(word)}
      className="ml-4 inline-flex h-12 w-12 items-center justify-center rounded-full border-2 border-stone-200 bg-stone-100 text-stone-600 shadow-sm transition hover:scale-105 hover:bg-white"
      aria-label="play pronunciation"
    >
      <Volume2 className="h-5 w-5" />
    </button>
  );
}

function Progress({ stageIndex, wordIndex, total }) {
  return (
    <div className="mb-8">
      <div className="mb-3 flex items-center justify-between text-sm text-stone-500">
        <span>Word {Math.min(wordIndex + 1, total)} / {total}</span>
        <span>Step {stageIndex + 1} / {STAGES.length}</span>
      </div>
      <div className="grid grid-cols-4 gap-2">
        {STAGES.map((stage, index) => (
          <div key={stage.key} className="h-2 overflow-hidden rounded-full bg-stone-200">
            <div
              className={`h-full rounded-full transition-all duration-500 ${index <= stageIndex ? "bg-stone-950" : "bg-transparent"}`}
              style={{ width: index <= stageIndex ? "100%" : "0%" }}
            />
          </div>
        ))}
      </div>
    </div>
  );
}

function WordHeader({ word, mode = "large" }) {
  return (
    <div>
      <div className="flex items-center">
        <h1 className={`${mode === "large" ? "text-7xl md:text-8xl" : "text-5xl md:text-6xl"} font-serif font-black tracking-tight text-stone-950`}>
          {word.word}
        </h1>
        <AudioButton word={word.word} />
      </div>
      <p className="mt-5 font-mono text-2xl text-slate-600">{word.phonetic}</p>
    </div>
  );
}

function Feedback({ type, title, children }) {
  const isCorrect = type === "correct";
  return (
    <motion.div
      initial={{ opacity: 0, y: 12 }}
      animate={{ opacity: 1, y: 0 }}
      className={`mt-7 rounded-2xl border-l-4 p-6 ${isCorrect ? "border-green-600 bg-emerald-50 text-green-700" : "border-red-500 bg-red-50 text-red-700"}`}
    >
      <div className="mb-2 flex items-center gap-2 text-xl font-semibold">
        {isCorrect ? <Check className="h-5 w-5" /> : <X className="h-5 w-5" />}
        {title}
      </div>
      <p className="text-lg leading-relaxed">{children}</p>
    </motion.div>
  );
}

function TopBar({ mode, setMode, wrongWords, words, onUpload, uploadMessage }) {
  return (
    <div className="mb-8 rounded-3xl border border-stone-200 bg-white/60 p-4 shadow-sm backdrop-blur">
      <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
        <div>
          <h1 className="text-xl font-bold text-stone-950">Fifi Vocab Trainer</h1>
          <p className="mt-1 text-sm text-stone-500">语境理解 → 单词详情 → 听音写词 → 看义写词</p>
        </div>
        <div className="flex flex-wrap gap-3">
          <Button onClick={() => setMode("learn")} className={`rounded-2xl px-5 ${mode === "learn" ? "bg-stone-950" : "bg-stone-200 text-stone-900 hover:bg-stone-300"}`}>
            <BookOpen className="mr-2 h-4 w-4" /> 学习模式
          </Button>
          <Button onClick={() => setMode("review")} className={`rounded-2xl px-5 ${mode === "review" ? "bg-stone-950" : "bg-stone-200 text-stone-900 hover:bg-stone-300"}`}>
            <ListChecks className="mr-2 h-4 w-4" /> 复习模式 {wrongWords.length > 0 ? `(${wrongWords.length})` : ""}
          </Button>
          <label className="inline-flex cursor-pointer items-center rounded-2xl bg-green-600 px-5 py-2 text-sm font-medium text-white transition hover:bg-green-700">
            <Upload className="mr-2 h-4 w-4" /> 上传词库
            <input type="file" accept=".json,.csv,.xlsx,.xls" className="hidden" onChange={onUpload} />
          </label>
        </div>
      </div>
      <div className="mt-4 flex flex-col gap-2 text-sm text-stone-500 md:flex-row md:items-center md:justify-between">
        <p>当前词库：{words.length} 个词。支持上传 Excel / CSV / JSON。</p>
        {uploadMessage && <p className="font-medium text-green-700">{uploadMessage}</p>}
      </div>
    </div>
  );
}

function WrongBook({ wrongRecords, clearWrongBook, startReview }) {
  return (
    <Card className="mb-8 border-2 border-red-100 bg-red-50/60 shadow-none">
      <CardContent className="p-6">
        <div className="mb-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <AlertCircle className="h-6 w-6 text-red-600" />
            <div>
              <h2 className="text-xl font-bold text-stone-950">错题本</h2>
              <p className="text-sm text-stone-500">答错的单词会自动加入这里，复习模式会优先练这些词。</p>
            </div>
          </div>
          <div className="flex gap-2">
            <Button disabled={wrongRecords.length === 0} onClick={startReview} className="rounded-xl bg-red-600 hover:bg-red-700">开始复习</Button>
            <Button disabled={wrongRecords.length === 0} onClick={clearWrongBook} className="rounded-xl bg-stone-200 text-stone-900 hover:bg-stone-300">清空</Button>
          </div>
        </div>
        {wrongRecords.length === 0 ? (
          <p className="rounded-2xl bg-white/60 p-4 text-stone-500">目前还没有错题。学生答错后，这里会显示单词、错误环节和错误次数。</p>
        ) : (
          <div className="grid gap-3 md:grid-cols-2">
            {wrongRecords.map((item) => (
              <div key={item.word} className="rounded-2xl bg-white/70 p-4">
                <div className="flex items-center justify-between">
                  <p className="text-lg font-bold text-stone-950">{item.word}</p>
                  <span className="rounded-full bg-red-100 px-3 py-1 text-xs font-semibold text-red-700">错 {item.count} 次</span>
                </div>
                <p className="mt-1 text-sm text-stone-500">{item.meaning}</p>
                <p className="mt-2 text-xs text-stone-400">最近错误环节：{item.lastStage}</p>
              </div>
            ))}
          </div>
        )}
      </CardContent>
    </Card>
  );
}

function ContextStage({ word, onNext, markWrong }) {
  const [selected, setSelected] = useState(null);
  const isCorrect = selected === word.correctOption;

  function choose(option) {
    setSelected(option);
    if (option !== word.correctOption) markWrong(word, "语境理解");
  }

  return (
    <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
      <Pill tone="blue">L1 · 语境理解</Pill>
      <div className="mt-10"><WordHeader word={word} /></div>
      <p className="mt-9 text-xl text-slate-600">这个词在句子里是什么意思？</p>
      <Card className="mt-6 border-2 border-stone-200 bg-white/70 shadow-none">
        <CardContent className="p-8 font-serif text-2xl leading-relaxed md:text-3xl">{word.example}</CardContent>
      </Card>
      <div className="mt-8 grid grid-cols-1 gap-4 md:grid-cols-2">
        {word.options.map((option) => {
          const active = selected === option;
          const good = active && option === word.correctOption;
          const bad = active && option !== word.correctOption;
          return (
            <button key={option} onClick={() => choose(option)} className={`rounded-2xl border-2 bg-white/60 p-6 text-left text-xl transition hover:border-green-300 hover:bg-green-50 ${good ? "border-green-400 bg-emerald-50 text-green-700" : ""} ${bad ? "border-red-300 bg-red-50 text-red-600" : "border-stone-200 text-stone-900"}`}>{option}</button>
          );
        })}
      </div>
      {selected && <Feedback type={isCorrect ? "correct" : "wrong"} title={isCorrect ? "正确！" : "已加入错题本"}>{isCorrect ? word.explanation : `这个句子中需要表达“${word.correctOption}”的意思。`}</Feedback>}
      <div className="mt-8 flex items-center gap-4">
        <Button disabled={!isCorrect} onClick={onNext} className="rounded-2xl bg-stone-950 px-8 py-7 text-xl hover:bg-stone-800">下一步 <ArrowRight className="ml-2 h-5 w-5" /></Button>
        {selected && !isCorrect && <p className="text-stone-500">选对后进入单词详情</p>}
      </div>
    </motion.div>
  );
}

function DetailStage({ word, onNext }) {
  const [showHint, setShowHint] = useState(true);
  return (
    <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
      <Pill tone="blue">单词详情</Pill>
      <div className="mt-12"><WordHeader word={word} /></div>
      <div className="mt-12 space-y-8 text-slate-700">
        <p className="text-3xl"><span className="font-semibold text-stone-950">{word.partOfSpeech}</span>{word.meaning}</p>
        <div>
          <p className="mb-4 text-xl font-semibold text-slate-400">例句：</p>
          <Card className="border-2 border-stone-200 bg-white/70 shadow-none"><CardContent className="p-8 font-serif text-2xl leading-relaxed md:text-3xl">{word.example}</CardContent></Card>
        </div>
        <div className="rounded-2xl bg-stone-100 p-6">
          <button onClick={() => setShowHint(!showHint)} className="mb-3 flex items-center gap-2 font-semibold text-stone-700">{showHint ? <EyeOff className="h-5 w-5" /> : <Eye className="h-5 w-5" />}记忆提示</button>
          {showHint && <p className="text-xl leading-relaxed text-stone-700">{word.hint}</p>}
        </div>
      </div>
      <Button onClick={onNext} className="mt-10 rounded-2xl bg-stone-950 px-8 py-7 text-xl hover:bg-stone-800">继续练习 <ArrowRight className="ml-2 h-5 w-5" /></Button>
    </motion.div>
  );
}

function SpellingStage({ word, mode, onNext, markWrong }) {
  const [answer, setAnswer] = useState("");
  const [checked, setChecked] = useState(false);
  const isCorrect = answer.trim().toLowerCase() === word.word.toLowerCase();
  const letters = useMemo(() => word.word.split(""), [word.word]);

  function submit() {
    setChecked(true);
    if (answer.trim() && !isCorrect) markWrong(word, mode === "sound" ? "听音写词" : "看义写词");
  }

  return (
    <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }}>
      <Pill tone="green">{mode === "sound" ? "拼写 · 听音写词" : "拼写 · 看义写词"}</Pill>
      <div className="mt-16 text-center">
        <h2 className="text-3xl font-bold text-stone-950">{mode === "sound" ? "听发音，写出英文单词" : `${word.partOfSpeech}${word.shortMeaning}`}</h2>
        <p className="mt-4 text-xl text-slate-600">{mode === "sound" ? "点击音频按钮，可以重复听这个单词。" : "根据中文意思写出英文单词。"}</p>
        {mode === "sound" ? <div className="mt-10 flex justify-center"><button onClick={() => speak(word.word)} className="flex h-20 w-20 items-center justify-center rounded-full bg-stone-950 text-white shadow-lg transition hover:scale-105"><Volume2 className="h-8 w-8" /></button></div> : <p className="mt-10 text-2xl text-slate-500">{word.example}</p>}
        <div className="mx-auto mt-12 flex max-w-2xl justify-center gap-2">
          {letters.map((letter, index) => <div key={`${letter}-${index}`} className="w-12 border-b-4 border-green-600 pb-2 text-center font-mono text-3xl text-green-700">{answer[index] || ""}</div>)}
        </div>
        <input value={answer} onChange={(e) => { setAnswer(e.target.value); setChecked(false); }} onKeyDown={(e) => e.key === "Enter" && submit()} placeholder="在这里输入单词" className={`mx-auto mt-10 block w-full max-w-4xl rounded-2xl border-2 bg-emerald-50/60 px-8 py-6 text-center font-mono text-3xl outline-none transition ${checked && isCorrect ? "border-green-600" : "border-green-500"}`} />
        {checked && <Feedback type={isCorrect ? "correct" : "wrong"} title={isCorrect ? "拼写正确！" : "已加入错题本"}>{isCorrect ? word.word : `正确拼写是 ${word.word}，注意字母顺序。`}</Feedback>}
        <div className="mt-10 flex justify-start gap-4">
          <Button onClick={submit} className="rounded-2xl bg-stone-950 px-8 py-7 text-xl hover:bg-stone-800">检查答案</Button>
          <Button disabled={!checked || !isCorrect} onClick={onNext} className="rounded-2xl bg-stone-950 px-8 py-7 text-xl hover:bg-stone-800">下一步 <ArrowRight className="ml-2 h-5 w-5" /></Button>
        </div>
      </div>
    </motion.div>
  );
}

function CompleteStage({ onRestart, wordIndex, setWordIndex, total, mode, setMode }) {
  const hasNext = wordIndex + 1 < total;
  return (
    <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="text-center">
      <div className="mx-auto flex h-20 w-20 items-center justify-center rounded-full bg-emerald-100 text-green-700"><Check className="h-10 w-10" /></div>
      <h2 className="mt-8 text-5xl font-serif font-black text-stone-950">本轮完成！</h2>
      <p className="mx-auto mt-5 max-w-2xl text-xl leading-relaxed text-slate-600">你已经完成了“语境猜义 → 单词详情 → 听音写词 → 看义写词”的完整记忆闭环。</p>
      <div className="mt-10 flex justify-center gap-4">
        <Button onClick={onRestart} className="rounded-2xl bg-stone-950 px-8 py-7 text-xl hover:bg-stone-800"><RotateCcw className="mr-2 h-5 w-5" /> 重练当前词</Button>
        {hasNext ? <Button onClick={() => setWordIndex(wordIndex + 1)} className="rounded-2xl bg-green-600 px-8 py-7 text-xl hover:bg-green-700">下一个单词 <ArrowRight className="ml-2 h-5 w-5" /></Button> : <Button onClick={() => { setMode(mode === "review" ? "learn" : "review"); setWordIndex(0); }} className="rounded-2xl bg-green-600 px-8 py-7 text-xl hover:bg-green-700">切换模式 <ArrowRight className="ml-2 h-5 w-5" /></Button>}
      </div>
    </motion.div>
  );
}

function UploadGuide() {
  return (
    <details className="mb-8 rounded-3xl border border-stone-200 bg-white/50 p-5 text-stone-600">
      <summary className="cursor-pointer font-semibold text-stone-900">上传词库格式说明</summary>
      <div className="mt-4 grid gap-4 md:grid-cols-3">
        <div className="rounded-2xl bg-stone-100 p-4">
          <p className="mb-2 font-semibold text-stone-900">Excel 表头</p>
          <code className="block whitespace-pre-wrap text-sm">word | phonetic | partOfSpeech | meaning | shortMeaning | example | options | correctOption | explanation | hint</code>
          <p className="mt-2 text-sm">Excel 第一行必须是字段名。options 用 | 分隔，例如：手册|小册子|传单|指南</p>
        </div>
        <div className="rounded-2xl bg-stone-100 p-4">
          <p className="mb-2 font-semibold text-stone-900">CSV 表头</p>
          <code className="block whitespace-pre-wrap text-sm">word,phonetic,partOfSpeech,meaning,shortMeaning,example,options,correctOption,explanation,hint</code>
          <p className="mt-2 text-sm">CSV 和 Excel 用同一套字段。</p>
        </div>
        <div className="rounded-2xl bg-stone-100 p-4">
          <p className="mb-2 font-semibold text-stone-900">必填字段</p>
          <p className="text-sm leading-relaxed">最少只需要填 <b>word</b> 和 <b>meaning</b>。其他字段不填时，系统会自动补默认内容；但为了课堂效果，建议至少补 example 和 options。</p>
        </div>
      </div>
    </details>
  );
}

export default function VocabLearningApp() {
  const [words, setWords] = useState(() => {
    try {
      return JSON.parse(localStorage.getItem("fifi_vocab_words")) || SAMPLE_WORDS;
    } catch {
      return SAMPLE_WORDS;
    }
  });
  const [wrongBook, setWrongBook] = useState(() => {
    try {
      return JSON.parse(localStorage.getItem("fifi_vocab_wrong_book")) || {};
    } catch {
      return {};
    }
  });
  const [mode, setMode] = useState("learn");
  const [stageIndex, setStageIndex] = useState(0);
  const [wordIndex, setWordIndexRaw] = useState(0);
  const [uploadMessage, setUploadMessage] = useState("");
  const [role, setRole] = useState("student");
  const [studentName, setStudentName] = useState(() => localStorage.getItem("fifi_vocab_student_name") || "Mike");

  const wrongRecords = Object.values(wrongBook).sort((a, b) => b.count - a.count);
  const wrongWords = wrongRecords.map((record) => words.find((word) => word.word === record.word)).filter(Boolean);
  const activeWords = mode === "review" && wrongWords.length > 0 ? wrongWords : words;
  const word = activeWords[Math.min(wordIndex, activeWords.length - 1)];
  const stage = STAGES[stageIndex]?.key || "complete";

  React.useEffect(() => {
    localStorage.setItem("fifi_vocab_words", JSON.stringify(words));
  }, [words]);

  React.useEffect(() => {
    localStorage.setItem("fifi_vocab_wrong_book", JSON.stringify(wrongBook));
  }, [wrongBook]);

  React.useEffect(() => {
    localStorage.setItem("fifi_vocab_student_name", studentName);
  }, [studentName]);

  function setWordIndex(next) {
    setWordIndexRaw(typeof next === "function" ? next(wordIndex) : next);
    setStageIndex(0);
  }

  function setModeSafely(nextMode) {
    setMode(nextMode);
    setWordIndexRaw(0);
    setStageIndex(0);
  }

  function markWrong(word, lastStage) {
    setWrongBook((prev) => ({
      ...prev,
      [word.word]: {
        word: word.word,
        meaning: word.meaning,
        count: (prev[word.word]?.count || 0) + 1,
        lastStage,
      },
    }));
  }

  async function handleUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    let imported = [];
    try {
      if (file.name.endsWith(".json")) {
        const text = await file.text();
        const data = JSON.parse(text);
        imported = Array.isArray(data) ? data.map(normaliseWord).filter(Boolean) : [];
      } else if (file.name.endsWith(".xlsx") || file.name.endsWith(".xls")) {
        const arrayBuffer = await file.arrayBuffer();
        imported = parseExcel(arrayBuffer);
      } else {
        const text = await file.text();
        imported = parseCSV(text);
      }
    } catch (error) {
      setUploadMessage("上传失败：请检查 Excel / CSV / JSON 格式");
      return;
    }
    if (!imported.length) {
      setUploadMessage("没有识别到有效单词，请检查 word 和 meaning 字段");
      return;
    }
    setWords(imported);
    setWrongBook({});
    setModeSafely("learn");
    setUploadMessage(`已导入 ${imported.length} 个单词`);
    event.target.value = "";
  }

  const nextStage = () => setStageIndex((s) => Math.min(s + 1, STAGES.length));
  const restart = () => setStageIndex(0);

  return (
    <main className="min-h-screen bg-[#f3efe7] px-6 py-8 text-stone-950 md:px-12">
      <div className="mx-auto max-w-6xl">
        <div className="mb-6 flex flex-col gap-3 rounded-3xl border border-stone-200 bg-white/70 p-4 md:flex-row md:items-center md:justify-between">
          <div>
            <p className="text-sm font-semibold text-stone-500">当前学生</p>
            <input value={studentName} onChange={(e) => setStudentName(e.target.value)} className="mt-1 rounded-xl border border-stone-200 bg-white px-4 py-2 text-lg font-bold outline-none" />
          </div>
          <div className="flex gap-2">
            <Button onClick={() => setRole("student")} className={`rounded-2xl ${role === "student" ? "bg-stone-950" : "bg-stone-200 text-stone-900 hover:bg-stone-300"}`}>学生端</Button>
            <Button onClick={() => setRole("teacher")} className={`rounded-2xl ${role === "teacher" ? "bg-stone-950" : "bg-stone-200 text-stone-900 hover:bg-stone-300"}`}>老师后台</Button>
          </div>
        </div>
        <TopBar mode={mode} setMode={setModeSafely} wrongWords={wrongWords} words={words} onUpload={handleUpload} uploadMessage={uploadMessage} />
        <UploadGuide />
        {role === "teacher" && (
          <Card className="mb-8 border-2 border-stone-200 bg-white/70 shadow-none">
            <CardContent className="p-6">
              <h2 className="text-2xl font-bold text-stone-950">老师后台 · 学习数据</h2>
              <p className="mt-2 text-stone-500">这里是给老师看的简易数据面板。正式版可以换成学生账号和云端数据库。</p>
              <div className="mt-6 grid gap-4 md:grid-cols-4">
                <div className="rounded-2xl bg-stone-100 p-5"><p className="text-sm text-stone-500">学生</p><p className="mt-2 text-2xl font-bold">{studentName}</p></div>
                <div className="rounded-2xl bg-stone-100 p-5"><p className="text-sm text-stone-500">当前词库</p><p className="mt-2 text-2xl font-bold">{words.length} 个</p></div>
                <div className="rounded-2xl bg-stone-100 p-5"><p className="text-sm text-stone-500">错题数量</p><p className="mt-2 text-2xl font-bold">{wrongRecords.length} 个</p></div>
                <div className="rounded-2xl bg-stone-100 p-5"><p className="text-sm text-stone-500">错误总次数</p><p className="mt-2 text-2xl font-bold">{wrongRecords.reduce((sum, item) => sum + item.count, 0)} 次</p></div>
              </div>
            </CardContent>
          </Card>
        )}
        <WrongBook wrongRecords={wrongRecords} clearWrongBook={() => setWrongBook({})} startReview={() => setModeSafely("review")} />
        {mode === "review" && wrongWords.length === 0 ? (
          <Card className="border-2 border-stone-200 bg-white/70 shadow-none"><CardContent className="p-10 text-center"><h2 className="text-3xl font-bold">现在还没有错题</h2><p className="mt-3 text-stone-500">先进入学习模式，答错的单词会自动进入复习模式。</p><Button onClick={() => setModeSafely("learn")} className="mt-6 rounded-2xl bg-stone-950 px-8 py-6 text-lg">返回学习模式</Button></CardContent></Card>
        ) : (
          <>
            <Progress stageIndex={Math.min(stageIndex, STAGES.length - 1)} wordIndex={wordIndex} total={activeWords.length} />
            <AnimatePresence mode="wait">
              <motion.div key={`${word.word}-${stageIndex}-${mode}`} initial={{ opacity: 0, y: 14 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -14 }} transition={{ duration: 0.24 }} className="rounded-[2rem] bg-[#f3efe7]">
                {stage === "context" && <ContextStage word={word} onNext={nextStage} markWrong={markWrong} />}
                {stage === "detail" && <DetailStage word={word} onNext={nextStage} />}
                {stage === "sound" && <SpellingStage word={word} mode="sound" onNext={nextStage} markWrong={markWrong} />}
                {stage === "meaning" && <SpellingStage word={word} mode="meaning" onNext={nextStage} markWrong={markWrong} />}
                {stage === "complete" && <CompleteStage onRestart={restart} wordIndex={wordIndex} setWordIndex={setWordIndex} total={activeWords.length} mode={mode} setMode={setModeSafely} />}
              </motion.div>
            </AnimatePresence>
          </>
        )}
      </div>
    </main>
  );
}
