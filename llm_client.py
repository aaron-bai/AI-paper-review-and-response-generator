"""
llm_client.py
Lightweight HTTP client for OpenAI-compatible LLM APIs.
"""

import json
import urllib.request
import urllib.error
from typing import Optional


class LLMClient:
    """Send chat-completion requests to an OpenAI-compatible endpoint."""

    def __init__(
        self,
        api_url: str,
        api_token: str,
        model: str,
        temperature: float = 0.7,
    ):
        self.api_url = api_url.rstrip("/")
        self.api_token = api_token
        self.model = model
        self.temperature = temperature

    def chat(
        self,
        system_prompt: str,
        user_prompt: str,
        max_tokens: Optional[int] = 64000,
    ) -> str:
        """
        Send a chat completion request and return the assistant reply as a string.

        Parameters
        ----------
        system_prompt : str
            The system message that sets the assistant's behavior.
        user_prompt : str
            The user message (the actual query).
        max_tokens : int, optional
            Maximum number of tokens to generate.

        Returns
        -------
        str
            The content of the first assistant choice.
        """
        payload: dict = {
            "model": self.model,
            "temperature": self.temperature,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        }
        if max_tokens is not None:
            payload["max_tokens"] = max_tokens

        data = json.dumps(payload).encode("utf-8")
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_token}",
        }

        endpoint = f"{self.api_url}/chat/completions"
        request = urllib.request.Request(endpoint, data=data, headers=headers, method="POST")

        try:
            with urllib.request.urlopen(request) as response:
                body = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            error_body = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(
                f"LLM API request failed with status {exc.code}: {error_body}"
            ) from exc

        try:
            return body["choices"][0]["message"]["content"]
        except (KeyError, IndexError) as exc:
            raise RuntimeError(
                f"Unexpected LLM API response format: {body}"
            ) from exc
