"""
Market Research Crew
====================
Orchestrates three AI agents working together to produce
competitive intelligence reports.

Flow: Scout (research) → Analyst (insights) → Reporter (final report)

Supports FREE LLM providers: Google Gemini, Groq, Ollama
"""

import os
from datetime import datetime

from crewai import Crew, LLM, Process
from crewai_tools import SerperDevTool, ScrapeWebsiteTool

from .agents import create_scout_agent, create_analyst_agent, create_reporter_agent
from .tasks import create_research_task, create_analysis_task, create_report_task
from .tools.ppt_generator import generate_pptx


def setup_llm() -> LLM:
    """Configure the LLM based on available API keys (all FREE options)."""
    model_name = os.getenv("MODEL_NAME", "")

    # Option 1: Google Gemini (FREE)
    if os.getenv("GEMINI_API_KEY"):
        model = model_name or "gemini/gemini-2.0-flash"
        print(f"  LLM: Google Gemini ({model})")
        return LLM(
            model=model,
            api_key=os.getenv("GEMINI_API_KEY"),
        )

    # Option 2: Groq (FREE)
    if os.getenv("GROQ_API_KEY"):
        model = model_name or "groq/llama-3.3-70b-versatile"
        print(f"  LLM: Groq ({model})")
        return LLM(
            model=model,
            api_key=os.getenv("GROQ_API_KEY"),
        )

    # Option 3: Ollama (FREE - local)
    if model_name and model_name.startswith("ollama/"):
        print(f"  LLM: Ollama Local ({model_name})")
        return LLM(
            model=model_name,
            base_url="http://localhost:11434",
        )

    # Option 4: OpenAI (paid, if user has key)
    if os.getenv("OPENAI_API_KEY"):
        model = model_name or "gpt-4o-mini"
        print(f"  LLM: OpenAI ({model})")
        return LLM(model=model)

    return None


class MarketResearchCrew:
    """Manages the multi-agent market research workflow."""

    def __init__(self):
        self.llm = setup_llm()
        self.tools = self._setup_tools()

    def _setup_tools(self) -> dict:
        """Initialize tools available to agents."""
        tools = {}

        # Web search tool (requires SERPER_API_KEY - free 2,500 searches)
        if os.getenv("SERPER_API_KEY"):
            tools["search"] = SerperDevTool()
            tools["scraper"] = ScrapeWebsiteTool()
        else:
            print(
                "  WARNING: No SERPER_API_KEY found.\n"
                "  Scout agent will work with limited capabilities.\n"
                "  Get a free key at https://serper.dev"
            )

        return tools

    def run(self, company: str, competitors: list[str]) -> str:
        """
        Execute the full market research pipeline.

        Args:
            company: The client company name
            competitors: List of competitor names to research

        Returns:
            The generated report content
        """
        print("\n" + "=" * 60)
        print("  MARKET RESEARCH AGENT - Starting Analysis")
        print(f"  Client: {company}")
        print(f"  Competitors: {', '.join(competitors)}")
        print("=" * 60 + "\n")

        # Generate output filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = company.lower().replace(" ", "_")
        output_file = f"output/report_{safe_name}_{timestamp}.md"

        # Prepare tool lists for each agent
        # Note: Groq/free models have issues with tool calling,
        # so we only pass tools when using OpenAI or paid providers
        model_name = os.getenv("MODEL_NAME", "")
        use_tools = not model_name.startswith(("groq/", "ollama/"))
        scout_tools = list(self.tools.values()) if use_tools else []
        analyst_tools = []  # Analyst works with data from Scout
        reporter_tools = []  # Reporter works with analysis data

        if not use_tools and self.tools:
            print("  Note: Using LLM knowledge (no web search) for compatibility")
            print("        Add OpenAI key for live web search capability\n")

        # Create agents with the configured LLM
        print("[1/3] Deploying Scout Agent...")
        scout = create_scout_agent(scout_tools, self.llm)

        print("[2/3] Deploying Analyst Agent...")
        analyst = create_analyst_agent(analyst_tools, self.llm)

        print("[3/3] Deploying Reporter Agent...")
        reporter = create_reporter_agent(reporter_tools, self.llm)

        # Create tasks (sequential pipeline)
        research_task = create_research_task(scout, company, competitors)
        analysis_task = create_analysis_task(analyst, company, competitors)
        report_task = create_report_task(reporter, company, competitors, output_file)

        # Assemble the crew
        crew = Crew(
            agents=[scout, analyst, reporter],
            tasks=[research_task, analysis_task, report_task],
            process=Process.sequential,  # Tasks run in order
            verbose=True,
            max_rpm=4,  # Rate limit to stay within Groq free tier (12k TPM)
        )

        # Execute the research pipeline
        print("\n" + "-" * 60)
        print("  Crew assembled! Starting research pipeline...")
        print("  Scout -> Analyst -> Reporter")
        print("-" * 60 + "\n")

        result = crew.kickoff()
        report_text = str(result)

        # Generate PowerPoint presentation
        pptx_file = output_file.replace(".md", ".pptx")
        print("\n[PPT] Generating PowerPoint presentation...")
        try:
            generate_pptx(report_text, company, competitors, pptx_file)
            print(f"[PPT] Saved to: {pptx_file}")
        except Exception as e:
            print(f"[PPT] Warning: Could not generate PPTX: {e}")
            pptx_file = None

        print("\n" + "=" * 60)
        print("  RESEARCH COMPLETE!")
        print(f"  Markdown: {output_file}")
        if pptx_file:
            print(f"  PowerPoint: {pptx_file}")
        print("=" * 60 + "\n")

        return report_text
