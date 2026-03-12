"""
review_generator.py
Uses an LLM to generate peer-review questions for a research paper.

Only problems / weaknesses are raised — no praise or positive comments.
Each issue must reference the specific part of the paper it concerns.
"""

from llm_client import LLMClient


_SYSTEM_PROMPT = (
    "You are a rigorous and experienced peer reviewer for academic journals in the field "
    "of {discipline}. "
    "Your role is to critically evaluate research papers and identify their weaknesses, "
    "gaps, and problems. "
    "You must ONLY raise concerns, questions, and problems — do NOT mention any strengths "
    "or positive aspects of the paper. "
    "For every issue you raise, you must cite the specific section, paragraph, figure, or "
    "table in the original paper that the issue relates to. "
    "Format your output as a numbered list. Each item must follow this structure:\n"
    "  [Reference: <section/page/paragraph>; <original_text>] <The question or problem>\n\n"
    "Be thorough, precise, and academically rigorous."
    "Do not raise too many minor issues; focus on the most important problems that would be relevant for an academic peer review."
    "You must generate reviews on the frontier of current academic standards of the focal paper's discipline. "
    "When raise issues, you should describe why and how they are problems in very detail, as if you were explaining to the paper's authors. And you can raise some reference to the current state of the art in the field, if relevant. "
    "And the each explaination should be 300-500 words or more, so you need to be very detailed and specific in your explaination. "
    "Remember: ONLY RAISE PROBLEMS — NO PRAISE."
    "You should generate reviews in {language}."
    "If the paper language is not {language}, you should cite the original text in the paper's language, but write your review questions in {language}."
    " Don't translate the paper text in your output, just reference it as is. "
)

_USER_TEMPLATE = (
    "Please review the following research paper from the field of {discipline} "
    "and raise all questions and problems you find. Remember: list ONLY problems — "
    "no positive comments.\n\n"
    "You should foucus on these aspects when reviewing the paper:\n"
    "{review_aspects}\n\n"
    "You should generate reviews in {language}."
    "--- PAPER TEXT ---\n"
    "{paper_text}\n"
    "--- END OF PAPER TEXT ---\n\n"
    "Peer-review questions (problems only):"
)

# Limit paper text sent to LLM to stay within context windows.
_MAX_CHARS = 24000


class ReviewGenerator:
    """Generate peer-review questions (problems only) for a research paper."""

    def __init__(self, llm_client: LLMClient):
        self.llm = llm_client

    def generate(self, paper_text: str, discipline: str, language: str = "English", review_aspects: str = "") -> str:
        """
        Return a numbered list of review questions / problems.

        Parameters
        ----------
        paper_text : str
            Full or partial text of the research paper.
        discipline : str
            The academic discipline of the paper (from DisciplineDetector).
        language : str
            The language in which to generate the review questions.
        review_aspects : str
            Specific aspects to focus on when generating the review questions.

        Returns
        -------
        str
            Numbered list of review questions, each anchored to a paper reference.
        """
        excerpt = paper_text[:_MAX_CHARS]
        system_prompt = _SYSTEM_PROMPT.format(discipline=discipline, language=language)
        user_prompt = _USER_TEMPLATE.format(
            discipline=discipline, paper_text=excerpt, language=language, review_aspects=review_aspects
        )
        return self.llm.chat(
            system_prompt=system_prompt,
            user_prompt=user_prompt,
        ).strip()
