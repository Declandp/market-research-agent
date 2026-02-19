"""
Market Research Task Definitions
================================
Each task defines specific work for an agent to complete.
Tasks are executed sequentially: Research → Analysis → Report
"""

from crewai import Task, Agent


def create_research_task(agent: Agent, company: str, competitors: list[str]) -> Task:
    """Task for the Scout Agent to gather competitor intelligence."""
    competitor_list = ", ".join(competitors)

    return Task(
        description=f"""
        Research competitors for "{company}".
        Competitors: {competitor_list}

        For EACH competitor, do 2 web searches max, then summarize:
        1. What they do, pricing, key features
        2. Strengths, weaknesses, recent news

        Keep searches focused. Do NOT do more than 2 searches per competitor.
        """,
        expected_output=f"""
        A research summary for each competitor ({competitor_list}) with:
        - Company overview
        - Products and pricing
        - Key strengths and weaknesses
        - Recent developments
        """,
        agent=agent,
    )


def create_analysis_task(agent: Agent, company: str, competitors: list[str]) -> Task:
    """Task for the Analyst Agent to extract strategic insights."""
    competitor_list = ", ".join(competitors)

    return Task(
        description=f"""
        Using the research data collected on competitors ({competitor_list}),
        perform a comprehensive strategic analysis for "{company}".

        **Analysis Framework:**

        1. **SWOT Analysis (for each competitor)**
           - Strengths: What are they doing well?
           - Weaknesses: Where are they falling short?
           - Opportunities: What gaps can our client exploit?
           - Threats: What competitive risks do they pose?

        2. **Competitive Comparison Matrix**
           - Compare all competitors across key dimensions:
             * Pricing (budget / mid-range / premium)
             * Product breadth (narrow / moderate / wide)
             * Market presence (emerging / established / dominant)
             * Innovation level (lagging / keeping pace / leading)
             * Customer satisfaction (low / medium / high)

        3. **Market Gaps & Opportunities**
           - What needs are competitors NOT addressing?
           - What customer complaints keep appearing?
           - Where is the market headed that competitors aren't?
           - What pricing gaps exist?

        4. **Threat Assessment**
           - Rank competitors by threat level (1-10)
           - Identify each competitor's most dangerous advantage
           - Predict likely competitive moves in next 6-12 months

        5. **Strategic Recommendations**
           - Top 3 opportunities for our client to differentiate
           - Top 3 threats to watch and mitigate
           - Quick wins vs. long-term strategic moves

        **Important:**
        - Base ALL analysis on the research data provided
        - Be specific — avoid generic strategy advice
        - Quantify findings where possible
        - Prioritize actionability over comprehensiveness
        """,
        expected_output=f"""
        A strategic analysis document containing:
        - Individual SWOT analysis for each competitor
        - Competitive comparison matrix with ratings
        - Identified market gaps and opportunities (ranked by potential)
        - Threat assessment with competitor rankings
        - Specific, actionable strategic recommendations
        """,
        agent=agent,
    )


def create_report_task(
    agent: Agent, company: str, competitors: list[str], output_file: str
) -> Task:
    """Task for the Reporter Agent to create the final intelligence report."""
    competitor_list = ", ".join(competitors)

    return Task(
        description=f"""
        Create a professional Competitive Intelligence Report for "{company}"
        based on the research and analysis provided.

        **Report Structure:**

        # Competitive Intelligence Report: {company}
        **Date:** [Current Date]
        **Competitors Analyzed:** {competitor_list}

        ## 1. Executive Summary (5-7 bullet points)
        - Key findings at a glance
        - Most critical threats and opportunities
        - Top recommendation

        ## 2. Market Landscape Overview
        - Industry context
        - Current market dynamics
        - Key trends shaping the competitive environment

        ## 3. Competitor Profiles
        For each competitor:
        - Company snapshot (overview, size, focus)
        - Products/services and pricing
        - Strengths and weaknesses
        - Recent developments
        - Threat level rating (1-10)

        ## 4. Competitive Analysis
        - Comparison matrix (table format)
        - Market positioning map
        - SWOT highlights

        ## 5. Opportunities & Threats
        - Top 5 opportunities (ranked)
        - Top 5 threats (ranked)
        - Market gaps identified

        ## 6. Strategic Recommendations
        - Immediate actions (next 30 days)
        - Short-term strategy (next 3 months)
        - Long-term positioning (next 12 months)

        ## 7. Appendix
        - Data sources
        - Methodology notes
        - Confidence ratings for key findings

        **Formatting Requirements:**
        - Use markdown formatting throughout
        - Include tables where appropriate
        - Use bullet points for clarity
        - Bold key findings and recommendations
        - Keep language professional but accessible
        - Total length: 1,500-3,000 words
        """,
        expected_output=f"""
        A complete, professionally formatted Competitive Intelligence Report
        in markdown format, containing all sections listed above, with specific
        data, analysis, and actionable recommendations for {company}.
        """,
        agent=agent,
        output_file=output_file,
    )
