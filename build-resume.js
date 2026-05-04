const {
  Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat,
  BorderStyle, TabStopType, TabStopPosition, ExternalHyperlink, UnderlineType,
} = require('docx');
const fs = require('fs');

const ACCENT = "2E4057";
const RULE   = "4A90D9";
const BODY   = "1A1A1A";
const MUTED  = "555555";
const pt = n => n * 2;
const inch = n => Math.round(n * 1440);

function rule() {
  return new Paragraph({
    spacing: { before: 40, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: RULE, space: 1 } },
    children: [],
  });
}

function sectionHead(text) {
  return [
    new Paragraph({
      spacing: { before: 160, after: 20 },
      children: [new TextRun({
        text: text.toUpperCase(),
        bold: true, size: pt(11), color: ACCENT, font: "Calibri", characterSpacing: 60,
      })],
    }),
    rule(),
  ];
}

function jobHeader(title, company, dates) {
  return new Paragraph({
    spacing: { before: 140, after: 24 },
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    children: [
      new TextRun({ text: title, bold: true, size: pt(10.5), color: BODY, font: "Calibri" }),
      new TextRun({ text: "  ·  ", size: pt(10.5), color: MUTED, font: "Calibri" }),
      new TextRun({ text: company, size: pt(10.5), color: MUTED, italics: true, font: "Calibri" }),
      new TextRun({ text: "\t", font: "Calibri" }),
      new TextRun({ text: dates, size: pt(9.5), color: MUTED, font: "Calibri" }),
    ],
  });
}

function bullet(text, sub) {
  const children = [new TextRun({ text, size: pt(10), color: BODY, font: "Calibri" })];
  if (sub) children.push(new TextRun({ text: "  " + sub, size: pt(9.5), color: MUTED, italics: true, font: "Calibri" }));
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 28, after: 28 },
    children,
  });
}

function skillRow(label, value) {
  return new Paragraph({
    spacing: { before: 38, after: 38 },
    children: [
      new TextRun({ text: label + ":  ", bold: true, size: pt(10), color: ACCENT, font: "Calibri" }),
      new TextRun({ text: value, size: pt(10), color: BODY, font: "Calibri" }),
    ],
  });
}

function projectHead(name, url, stack, dates) {
  return [
    new Paragraph({
      spacing: { before: 130, after: 20 },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [
        new TextRun({ text: name, bold: true, size: pt(10.5), color: BODY, font: "Calibri" }),
        new TextRun({ text: "\t", font: "Calibri" }),
        new ExternalHyperlink({
          link: url,
          children: [new TextRun({
            text: url.replace(/^https?:\/\//, ''),
            size: pt(9.5), color: RULE,
            underline: { type: UnderlineType.SINGLE, color: RULE },
            font: "Calibri",
          })],
        }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: 28 },
      children: [new TextRun({ text: stack, size: pt(9.5), color: MUTED, italics: true, font: "Calibri" })],
    }),
  ];
}

function trainingGroup(label, courses) {
  return new Paragraph({
    spacing: { before: 44, after: 44 },
    children: [
      new TextRun({ text: label + ":  ", bold: true, size: pt(9.5), color: ACCENT, font: "Calibri" }),
      new TextRun({ text: courses.join("  ·  "), size: pt(9.5), color: BODY, font: "Calibri" }),
    ],
  });
}

function eduEntry(degree, school, dates, gpa) {
  return [
    new Paragraph({
      spacing: { before: 100, after: 20 },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [
        new TextRun({ text: degree, bold: true, size: pt(10.5), color: BODY, font: "Calibri" }),
        new TextRun({ text: "\t", font: "Calibri" }),
        new TextRun({ text: dates, size: pt(9.5), color: MUTED, font: "Calibri" }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      children: [
        new TextRun({ text: school, size: pt(9.5), color: MUTED, italics: true, font: "Calibri" }),
        ...(gpa ? [new TextRun({ text: "  ·  GPA: " + gpa, size: pt(9.5), color: MUTED, font: "Calibri" })] : []),
      ],
    }),
  ];
}

// ── Document ──────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: inch(0.25), hanging: inch(0.15) } } },
      }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: inch(0.65), bottom: inch(0.65), left: inch(0.75), right: inch(0.75) },
      },
    },
    children: [

      // ── Name ──
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "MICHAEL VALENTE", bold: true, size: pt(22), color: ACCENT, font: "Calibri", characterSpacing: 80 })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "Software Engineer  ·  AI / ML  ·  Algorithmic Trading  ·  Full-Stack", size: pt(11), color: MUTED, italics: true, font: "Calibri" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
        children: [
          new TextRun({ text: "Michael.valente1@gmail.com", size: pt(10), color: BODY, font: "Calibri" }),
          new TextRun({ text: "  |  508-948-8489", size: pt(10), color: BODY, font: "Calibri" }),
          new TextRun({ text: "  |  ", size: pt(10), color: MUTED, font: "Calibri" }),
          new ExternalHyperlink({ link: "https://github.com/mvalente27", children: [new TextRun({ text: "github.com/mvalente27", size: pt(10), color: RULE, underline: { type: UnderlineType.SINGLE, color: RULE }, font: "Calibri" })] }),
          new TextRun({ text: "  |  ", size: pt(10), color: MUTED, font: "Calibri" }),
          new ExternalHyperlink({ link: "https://www.linkedin.com/in/michael-v-91039634/", children: [new TextRun({ text: "linkedin.com/in/michael-v-91039634", size: pt(10), color: RULE, underline: { type: UnderlineType.SINGLE, color: RULE }, font: "Calibri" })] }),
          new TextRun({ text: "  |  ", size: pt(10), color: MUTED, font: "Calibri" }),
          new ExternalHyperlink({ link: "https://mvalente27.github.io", children: [new TextRun({ text: "mvalente27.github.io", size: pt(10), color: RULE, underline: { type: UnderlineType.SINGLE, color: RULE }, font: "Calibri" })] }),
        ],
      }),
      rule(),

      // ── Summary ──
      ...sectionHead("Summary"),
      new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [new TextRun({
          text: "Self-taught software engineer with 20+ years of executive leadership experience, now building production AI systems from the ground up. Author of 76,000+ lines of Python across four independent platforms: an IBKR algorithmic futures trading bot (Double DQN, Dueling architecture), a 41-module AI stock signal platform with a live FastAPI server and Android app, a DraftKings lineup optimizer (LightGBM + XGBoost + PyTorch + OR-Tools), and two live SaaS products. Combines deep business operations expertise with hands-on ML engineering and full-stack development. Completing an MBA in Data Analytics; 45+ DataCamp courses across Python, SQL, R, and AI/LLM engineering.",
          size: pt(10), color: BODY, font: "Calibri",
        })],
      }),

      // ── Technical Skills ──
      ...sectionHead("Technical Skills"),
      skillRow("Languages", "Python · TypeScript · JavaScript · SQL · R"),
      skillRow("AI / ML", "PyTorch · LightGBM · XGBoost · Double DQN · Reinforcement Learning · OpenAI API · Anthropic Claude API · Hugging Face · LLMOps · Prompt Engineering"),
      skillRow("Optimization", "OR-Tools (Google) · Kelly Criterion · Walk-forward backtesting"),
      skillRow("Frameworks", "FastAPI · React · Next.js · React Native · Node.js · Capacitor 6 (Android)"),
      skillRow("Data & APIs", "pandas · PostgreSQL · Supabase · yfinance · SEC EDGAR · FRED · ib_insync (IBKR) · Power BI · Tableau"),
      skillRow("Tools", "Git · GitHub Actions · Cloudflare Pages · Render · aiohttp · Pydantic · Typer"),

      // ── Software Projects ──
      ...sectionHead("Software Projects"),

      // Live-or-Die Bot
      ...projectHead(
        "Live-or-Die Bot — IBKR Algorithmic Futures Trading System",
        "https://github.com/mvalente27/Live-or-die-bot",
        "Python · PyTorch · Double DQN · ib_insync · IBKR TWS/Gateway · aiohttp"
      ),
      bullet("Built an autonomous futures trading bot (NQ, MNQ, YM, MYM, CL, GC, MBT) connected to Interactive Brokers via ib_insync; 60-second decision loop with 14-dimension state vector"),
      bullet("Agent architecture: Double DQN with Dueling network (separate value + advantage streams), soft Polyak target update (τ=0.005), prioritized replay, Huber loss, gradient clipping (norm 1.0), epsilon decay 0.15→0.01"),
      bullet("Indicator stack: Mark Fisher ACD (opening range, A/C levels, Fibonacci extensions), Elliott Wave v2 (rule-enforced W2/W3/W4 with Fibonacci targets), Williams Alligator, ADX, VWAP, market regime classifier"),
      bullet("Reward engine blends normalised PnL, rolling Sharpe ratio (60-step window), quadratic drawdown penalty, and a confluence bonus when ACD + Elliott + Alligator signals simultaneously agree"),
      bullet("Risk controls: daily loss circuit breaker, rolling drawdown gate (G4/G6), Kelly / proportional position sizing, auto-tier Micro vs Full contract resolution by account net liquidation"),

      // Q-Signals
      ...projectHead(
        "Q-Signals — AI-Powered Market Sentiment Research Platform",
        "https://q-signals.com",
        "Python · PyTorch · FastAPI · Supabase · Stripe · Capacitor 6 · Android · HTML / CSS / JS"
      ),
      bullet("Built an algorithmic sentiment research platform rating stocks and futures bullish / bearish / neutral — structured as a research publisher, not investment advice"),
      bullet("Dual-Gate Consensus: rating published only when |composite score| >= +/-0.10 AND 60%+ of 41 modules agree — eliminates false positives from any single noisy indicator"),
      bullet("41 modules span Elliott Wave, Fibonacci, 61 TA-Lib candlestick patterns, SEC EDGAR insider/institutional filings, FRED macro data, Reddit/StockTwits sentiment, and earnings surprise"),
      bullet("DQN agent (18K parameters, PyTorch) evaluates sentiment accuracy against T1/T2/T3 ATR-scaled price targets across 24h / 2wk / 3mo horizons; module weights adapt by hit rate"),
      bullet("FastAPI server (api.q-signals.com) with rate limiting, API key auth, CORS hardening; Stripe subscription billing; Supabase persistence; Android War Room app via Capacitor 6"),

      // DFS AI Agent
      ...projectHead(
        "DFS AI Agent — DraftKings Lineup Optimizer",
        "https://github.com/mvalente27/Live-or-die-bot",
        "Python · LightGBM · XGBoost · PyTorch · OR-Tools · FastAPI · Pydantic · Typer"
      ),
      bullet("Three-model ML ensemble (LightGBM, XGBoost, PyTorch 3×128 NN) blended by inverse-MAE adaptive weighting, with separate model registries per sport (NFL/NBA/MLB) and contest type"),
      bullet("OR-Tools constraint solver enforces $50K salary cap, 60% max player exposure, lineup cosine uniqueness ≥ 0.35, and position-specific rules; generates 50+ unique lineups per slate"),
      bullet("20-metric feature vector per player: Vegas implied totals, pace, usage, ownership leverage, social sentiment velocity, salary inefficiency, field simulation, and Kelly edge contribution"),
      bullet("Field simulator runs 10,000+ crowd lineups; walk-forward backtester with 2-season lookback and 30-day rolling validation windows; Kelly Criterion (quarter-Kelly) bankroll management"),
      bullet("FastAPI server with full CRUD API + Typer CLI (fetch-slate, run-slate, backtest, train, upload-csv commands)"),

      // UpkeepIQ
      ...projectHead(
        "UpkeepIQ — Home Maintenance SaaS",
        "https://upkeepiq.com",
        "React · Node.js · PostgreSQL"
      ),
      bullet("Built and deployed a home maintenance SaaS tracking appliances, warranty data, and service history — live at upkeepiq.com"),
      bullet("Automated maintenance reminders, scheduling logic, service-provider management, and cost tracking with monthly / annual rollups"),

      // ChoreQuest
      ...projectHead(
        "ChoreQuest — Gamified Family Task Manager",
        "https://github.com/mvalente27/ChoreQuest",
        "React Native · Supabase · TypeScript · Android (Google Play)"
      ),
      bullet("Built a family-focused Android app that turns household chores into quests: parents assign tasks and custom rewards, kids earn coins redeemable for screen time, treats, or activities"),
      bullet("XP system, level progression, household leaderboard, streak tracking, and recurring-task templates on a Supabase backend; beta-tested on Google Play with live users"),

      // ── Professional Experience ──
      ...sectionHead("Professional Experience"),

      jobHeader("Senior Vice President of Condominiums", "Brigs LLC  ·  Boston, MA", "Jan 2024 – Present"),
      bullet("Spearheaded data-driven retention initiatives — built KPI tracking dashboards that surfaced behavioral patterns, contributing to a 15% increase in client retention"),
      bullet("Partnered with technology vendors to upgrade internal systems, delivering a 25% gain in operational efficiency"),
      bullet("Secured and managed over $20M in capital project funding; required multi-year financial modeling, budget forecasting, and stakeholder reporting"),
      bullet("Led cross-departmental teams across a 40+ person organization, aligning strategy and execution at scale"),

      jobHeader("Vice President of Condominiums", "Brigs LLC  ·  Boston, MA", "Jan 2019 – Dec 2023"),
      bullet("Directed a team of 40+ professionals; implemented KPI reporting systems that improved performance transparency by 30%"),
      bullet("Developed and standardized departmental playbooks, reducing new-hire training time by 25%"),

      jobHeader("Senior Property Manager", "Great North Property Management  ·  Boston, MA", "Aug 2016 – Jan 2019"),
      bullet("Managed 15 associations (750+ units) with a 98% client retention rate; developed and presented annual operating and capital budgets exceeding $3M"),

      // ── Education ──
      ...sectionHead("Education"),
      ...eduEntry("Master of Business Administration (MBA), Data Analytics", "Johnson & Wales University, Providence, RI", "Oct 2023 – Aug 2025", "3.5"),
      ...eduEntry("Bachelor of Business Administration (BBA), International Business", "Southern New Hampshire University, Manchester, NH", "2022", "3.65"),

      // ── Technical Training ──
      ...sectionHead("Technical Training  (DataCamp — 45+ Courses)"),
      trainingGroup("AI & LLM Engineering",
        ["Developing AI Systems with the OpenAI API", "Working with the OpenAI API", "Prompt Engineering with the OpenAI API", "LLMOps Concepts", "Working with Hugging Face"]),
      trainingGroup("Python & Data Engineering",
        ["Introduction to Python", "Intermediate Python", "Python Toolbox", "Introduction to Functions in Python", "Data Manipulation with pandas", "Introduction to Data Visualization with Matplotlib", "Understanding Data Engineering"]),
      trainingGroup("SQL & Databases",
        ["Introduction to SQL", "Intermediate SQL", "Joining Data in SQL", "Introduction to Relational Databases in SQL", "Data Manipulation in SQL", "Exploratory Data Analysis in SQL", "Data-Driven Decision Making in SQL", "Applying SQL to Real-World Problems", "PostgreSQL Summary Stats and Window Functions"]),
      trainingGroup("R & Statistics",
        ["Introduction to R", "Intermediate R", "Introduction to the Tidyverse", "Data Manipulation with dplyr", "Introduction to Data Visualization with ggplot2", "Intermediate Data Visualization with ggplot2", "Exploratory Data Analysis in R", "Cleaning Data in R", "Dealing With Missing Data in R", "Introduction to Statistics in R", "Foundations of Probability in R", "Hypothesis Testing in R", "Inference for Categorical Data in R"]),
      trainingGroup("BI & Visualization",
        ["Introduction to Power BI", "Introduction to DAX in Power BI", "Data Visualization in Power BI", "Introduction to Tableau", "Creating Dashboards in Tableau", "Data Visualization in Tableau", "Introduction to Excel", "Data Visualization in Excel"]),
      trainingGroup("Cloud", ["Understanding Cloud Computing"]),

    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("resume-michael-valente.docx", buf);
  fs.copyFileSync("resume-michael-valente.docx", "D:/MV Items/resume-michael-valente.docx");
  console.log("Done — saved to portfolio-site/ and D:\\MV Items\\");
});
