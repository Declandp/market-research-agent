"""
Market Research Agent - CLI Entry Point
=======================================
Run competitive intelligence analysis using AI agents.

Usage:
    python run.py                          # Interactive mode
    python run.py --company "Acme Inc" --competitors "Rival1,Rival2,Rival3"

Examples:
    python run.py
    python run.py --company "Nike" --competitors "Adidas,Puma,New Balance"
    python run.py --company "Shopify" --competitors "WooCommerce,BigCommerce,Squarespace"
"""

import argparse
import os
import sys

from dotenv import load_dotenv


def get_inputs_interactive() -> tuple[str, list[str]]:
    """Get company and competitor names interactively."""
    print("\n" + "=" * 60)
    print("  AI MARKET RESEARCH AGENT")
    print("  Powered by CrewAI + Multi-Agent Intelligence")
    print("=" * 60)

    print("\nThis tool uses 3 AI agents to research your competitors:")
    print("  1. Scout Agent   - Gathers intelligence from the web")
    print("  2. Analyst Agent - Extracts strategic insights")
    print("  3. Reporter Agent - Creates professional reports")

    print("\n" + "-" * 60)
    company = input("\nYour company name: ").strip()
    if not company:
        print("Error: Company name is required.")
        sys.exit(1)

    print("\nEnter competitor names (comma-separated):")
    print("Example: Adidas, Puma, New Balance")
    competitor_input = input("\nCompetitors: ").strip()
    if not competitor_input:
        print("Error: At least one competitor is required.")
        sys.exit(1)

    competitors = [c.strip() for c in competitor_input.split(",") if c.strip()]

    print(f"\nResearching {len(competitors)} competitors for '{company}'...")
    print("This may take 3-5 minutes depending on the number of competitors.\n")

    return company, competitors


def get_inputs_cli(args) -> tuple[str, list[str]]:
    """Get inputs from command line arguments."""
    competitors = [c.strip() for c in args.competitors.split(",") if c.strip()]
    return args.company, competitors


def check_api_keys():
    """Verify required API keys are set (all FREE options)."""
    gemini_key = os.getenv("GEMINI_API_KEY")
    groq_key = os.getenv("GROQ_API_KEY")
    openai_key = os.getenv("OPENAI_API_KEY")
    model_name = os.getenv("MODEL_NAME", "")
    serper_key = os.getenv("SERPER_API_KEY")

    has_llm = gemini_key or groq_key or openai_key or model_name.startswith("ollama/")

    if not has_llm:
        print("\nERROR: No LLM API key found!")
        print("\nFREE setup options (choose one):")
        print()
        print("  Option 1: Google Gemini (RECOMMENDED - FREE)")
        print("    1. Go to https://aistudio.google.com/apikey")
        print("    2. Create a free API key")
        print("    3. Add to .env: GEMINI_API_KEY=your-key")
        print()
        print("  Option 2: Groq (FREE - very fast)")
        print("    1. Go to https://console.groq.com/keys")
        print("    2. Create a free API key")
        print("    3. Add to .env: GROQ_API_KEY=your-key")
        print()
        print("  Option 3: Ollama (FREE - runs locally)")
        print("    1. Install from https://ollama.com")
        print("    2. Run: ollama pull llama3.2")
        print("    3. Set in .env: MODEL_NAME=ollama/llama3.2")
        print()
        print("  First: cp .env.example .env")
        sys.exit(1)

    if not serper_key:
        print("  WARNING: No SERPER_API_KEY found.")
        print("  Scout agent needs this to search the web.")
        print("  Get a FREE key at: https://serper.dev (2,500 free searches)\n")

    if serper_key:
        print("  Web Search: Enabled (Serper)")
    print()


def main():
    # Load environment variables
    load_dotenv()

    # Parse arguments
    parser = argparse.ArgumentParser(
        description="AI-powered competitive intelligence research"
    )
    parser.add_argument("--company", type=str, help="Your company name")
    parser.add_argument(
        "--competitors", type=str, help="Comma-separated competitor names"
    )
    args = parser.parse_args()

    # Check API keys
    check_api_keys()

    # Get inputs
    if args.company and args.competitors:
        company, competitors = get_inputs_cli(args)
    else:
        company, competitors = get_inputs_interactive()

    # Run the research crew
    from src.crew import MarketResearchCrew

    crew = MarketResearchCrew()
    result = crew.run(company=company, competitors=competitors)

    print("\n" + "=" * 60)
    print("  FINAL REPORT")
    print("=" * 60)
    print(result)


if __name__ == "__main__":
    main()
