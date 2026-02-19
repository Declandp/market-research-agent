# AI Market Research Agent

> Multi-agent competitive intelligence system that automatically researches competitors and generates professional strategy reports.

## What It Does

This tool deploys **3 specialized AI agents** that work together like a research team:

```
Scout Agent          Analyst Agent         Reporter Agent
(Web Research)  →    (Strategic Analysis)  →  (Professional Report)
     |                      |                       |
Searches the web     Applies SWOT,          Creates executive-ready
for competitor       Porter's Five Forces,   intelligence report
data & news          and gap analysis        with recommendations
```

### Input
- Your company name
- List of competitors to research

### Output
- Professional competitive intelligence report (markdown)
- SWOT analysis for each competitor
- Competitive comparison matrix
- Strategic recommendations (30-day, 3-month, 12-month)

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Set Up API Keys

```bash
cp .env.example .env
# Edit .env and add your API keys
```

You'll need (ALL FREE):
- **Google Gemini API Key** (recommended) — [Get free key](https://aistudio.google.com/apikey)
- **Serper API Key** — [Get free key](https://serper.dev) (2,500 free searches)

Alternative LLM providers (also free):
- **Groq** — [Get free key](https://console.groq.com/keys)
- **Ollama** — [Install locally](https://ollama.com) (no API key needed)

### 3. Run the Agent

```bash
# Interactive mode
python run.py

# Or with arguments
python run.py --company "Your Company" --competitors "Rival1,Rival2,Rival3"
```

### Example

```bash
python run.py --company "Nike" --competitors "Adidas,Puma,New Balance"
```

Output: A complete competitive intelligence report saved to `output/`

## Project Structure

```
market-research-agent/
├── run.py                  # CLI entry point
├── src/
│   ├── crew.py             # Crew orchestration (coordinates agents)
│   ├── agents/
│   │   └── research_agents.py  # Agent definitions (Scout, Analyst, Reporter)
│   ├── tasks/
│   │   └── research_tasks.py   # Task definitions for each agent
│   └── tools/              # Custom tools (extensible)
├── config/                 # Configuration files
├── output/                 # Generated reports
│   └── sample_report.md    # Example output
├── requirements.txt
├── .env.example
└── README.md
```

## How It Works

### Agent Roles

| Agent | Role | What It Does |
|-------|------|--------------|
| **Scout** | Intelligence Gatherer | Searches the web for competitor data: products, pricing, news, reviews |
| **Analyst** | Strategy Expert | Applies SWOT analysis, identifies market gaps, ranks threats |
| **Reporter** | Report Writer | Creates professional reports with executive summaries and recommendations |

### Pipeline Flow

1. **Scout** searches the web for each competitor using Serper API
2. Scout passes research data to **Analyst**
3. Analyst applies strategic frameworks and extracts insights
4. Analyst passes analysis to **Reporter**
5. Reporter generates the final markdown report
6. Report is saved to `output/` directory

## Sample Output

See [output/sample_report.md](output/sample_report.md) for an example report.

## Tech Stack

- **Framework:** [CrewAI](https://www.crewai.com/) — Multi-agent orchestration
- **LLM:** Google Gemini / Groq / Ollama / OpenAI (configurable, free options available)
- **Search:** [Serper](https://serper.dev/) — Web search API (free tier)
- **Language:** Python 3.10+

## Customization

### Change the LLM

Edit `.env`:
```env
# Google Gemini (FREE - recommended)
GEMINI_API_KEY=your-key
MODEL_NAME=gemini/gemini-2.0-flash

# Groq (FREE - very fast)
GROQ_API_KEY=your-key
MODEL_NAME=groq/llama-3.3-70b-versatile

# Ollama (FREE - local, no key needed)
MODEL_NAME=ollama/llama3.2
```

### Add Custom Tools

Create new tools in `src/tools/` and add them to the agent in `src/crew.py`.

### Modify Report Format

Edit the report task in `src/tasks/research_tasks.py` to change the output structure.

## Cost Estimate

| Component | Cost |
|-----------|------|
| Google Gemini | **FREE** (1,500 requests/day) |
| Groq | **FREE** (rate limited) |
| Ollama (local) | **FREE** (unlimited) |
| Serper API | **FREE** (2,500 searches) |
| **Total per report** | **$0.00** |

## License

MIT

---

Built with CrewAI | Multi-Agent AI System
