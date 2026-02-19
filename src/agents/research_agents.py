"""
Market Research Agent Definitions
=================================
Three specialized agents that work together as a research crew:
1. Scout Agent - Gathers competitive intelligence from the web
2. Analyst Agent - Processes data and extracts strategic insights
3. Reporter Agent - Creates professional intelligence reports
"""

from crewai import Agent, LLM


def create_scout_agent(tools: list, llm: LLM = None) -> Agent:
    """The Scout: Finds and collects competitor data from the web."""
    kwargs = dict(
        role="Competitive Intelligence Scout",
        goal=(
            "Find comprehensive, accurate information about target competitors "
            "including their products, pricing, market positioning, recent news, "
            "social media presence, and customer sentiment."
        ),
        backstory=(
            "You are an elite competitive intelligence researcher with 15 years of "
            "experience at top consulting firms like McKinsey and BCG. You have a "
            "keen eye for finding hidden information and connecting dots that others "
            "miss. You're methodical, thorough, and never settle for surface-level "
            "data. You always verify information from multiple sources."
        ),
        tools=tools,
        verbose=True,
        allow_delegation=False,
        max_retry_limit=3,
    )
    if llm:
        kwargs["llm"] = llm
    return Agent(**kwargs)


def create_analyst_agent(tools: list, llm: LLM = None) -> Agent:
    """The Analyst: Processes competitor data and finds strategic insights."""
    kwargs = dict(
        role="Strategic Market Analyst",
        goal=(
            "Analyze competitor data to identify patterns, threats, opportunities, "
            "strengths, and weaknesses. Provide actionable strategic insights that "
            "give our client a competitive advantage."
        ),
        backstory=(
            "You are a senior data analyst with deep expertise in competitive "
            "strategy and market dynamics. You trained at Harvard Business School "
            "and spent a decade at Gartner analyzing market trends. You excel at "
            "turning raw data into clear strategic insights. You think in frameworks "
            "like SWOT, Porter's Five Forces, and Blue Ocean Strategy. You always "
            "back your analysis with evidence."
        ),
        tools=tools,
        verbose=True,
        allow_delegation=False,
        max_retry_limit=3,
    )
    if llm:
        kwargs["llm"] = llm
    return Agent(**kwargs)


def create_reporter_agent(tools: list, llm: LLM = None) -> Agent:
    """The Reporter: Creates polished intelligence reports."""
    kwargs = dict(
        role="Intelligence Report Specialist",
        goal=(
            "Create a comprehensive, professional competitive intelligence report "
            "that is clear, actionable, and visually organized. The report should "
            "be ready to present to C-level executives."
        ),
        backstory=(
            "You are an expert business writer who has created intelligence reports "
            "for Fortune 500 companies. You know how to distill complex analysis "
            "into clear, compelling narratives. Your reports are known for being "
            "both thorough and easy to act on. You always include an executive "
            "summary, key findings, and specific recommendations."
        ),
        tools=tools,
        verbose=True,
        allow_delegation=False,
        max_retry_limit=3,
    )
    if llm:
        kwargs["llm"] = llm
    return Agent(**kwargs)
