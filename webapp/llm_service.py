"""Multi-provider LLM abstraction layer.

Supports: OpenAI, Anthropic, DeepSeek, and custom OpenAI-compatible endpoints.
"""

from dataclasses import dataclass
from typing import Optional

from openai import OpenAI


@dataclass
class LLMResult:
    """Result from an LLM provider call."""
    text: str
    model: str
    provider: str


class LLMProvider:
    """Base class for LLM providers."""

    def __init__(self, api_key: str, model: str):
        self.api_key = api_key
        self.model = model

    def generate(
        self,
        system_prompt: str,
        user_message: str,
        temperature: float = 0.3,
        max_tokens: int = 16000,
    ) -> LLMResult:
        raise NotImplementedError


class OpenAIProvider(LLMProvider):
    """OpenAI provider with JSON mode support."""

    def __init__(self, api_key: str, model: str, base_url: Optional[str] = None):
        super().__init__(api_key, model)
        self.client = OpenAI(api_key=api_key, base_url=base_url)
        self.provider_name = "openai"

    def generate(
        self,
        system_prompt: str,
        user_message: str,
        temperature: float = 0.3,
        max_tokens: int = 16000,
    ) -> LLMResult:
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message},
            ],
            temperature=temperature,
            max_tokens=max_tokens,
            response_format={"type": "json_object"},
        )
        return LLMResult(
            text=response.choices[0].message.content or "",
            model=self.model,
            provider=self.provider_name,
        )


class DeepSeekProvider(OpenAIProvider):
    """DeepSeek provider — OpenAI-compatible with the same JSON mode."""

    def __init__(self, api_key: str, model: str):
        super().__init__(
            api_key=api_key,
            model=model,
            base_url="https://api.deepseek.com/v1",
        )
        self.provider_name = "deepseek"


class CustomProvider(OpenAIProvider):
    """Custom OpenAI-compatible provider with user-supplied base_url."""

    def __init__(self, api_key: str, model: str, base_url: str):
        super().__init__(api_key=api_key, model=model, base_url=base_url)
        self.provider_name = "custom"


class AnthropicProvider(LLMProvider):
    """Anthropic Claude provider.

    Claude doesn't have a native JSON mode, so we use the prefill technique:
    we include an assistant message starting with '{' to force JSON output.
    """

    def __init__(self, api_key: str, model: str):
        super().__init__(api_key, model)
        import anthropic
        self.client = anthropic.Anthropic(api_key=api_key)
        self.provider_name = "anthropic"

    def generate(
        self,
        system_prompt: str,
        user_message: str,
        temperature: float = 0.3,
        max_tokens: int = 32000,
    ) -> LLMResult:
        import anthropic

        # The prefill technique: prefill the assistant response with '{'
        # to force Claude to continue in JSON mode
        response = self.client.messages.create(
            model=self.model,
            max_tokens=max_tokens,
            temperature=temperature,
            system=system_prompt,
            messages=[
                {"role": "user", "content": user_message},
                # Prefill with '{' to force JSON output
                {"role": "assistant", "content": "{"},
            ],
        )

        text = response.content[0].text if response.content else ""
        # The response text from Claude will NOT include the prefill '{'
        # (Anthropic's API returns only the newly generated tokens)
        # But we add it back just in case the behavior varies
        full_text = "{" + text

        return LLMResult(
            text=full_text,
            model=self.model,
            provider=self.provider_name,
        )


def create_provider(
    provider_type: str,
    api_key: str,
    model: Optional[str] = None,
    base_url: Optional[str] = None,
) -> LLMProvider:
    """Factory function: create the appropriate LLM provider.

    Args:
        provider_type: "openai" | "anthropic" | "deepseek" | "custom"
        api_key: API key for the provider
        model: Optional model name override
        base_url: Required for "custom" type, optional otherwise

    Returns:
        Configured LLMProvider instance
    """
    from config import DEFAULT_MODELS

    if not model:
        model = DEFAULT_MODELS.get(provider_type, "")

    if provider_type == "openai":
        return OpenAIProvider(api_key=api_key, model=model)

    elif provider_type == "anthropic":
        return AnthropicProvider(api_key=api_key, model=model)

    elif provider_type == "deepseek":
        return DeepSeekProvider(api_key=api_key, model=model)

    elif provider_type == "custom":
        if not base_url:
            raise ValueError("base_url is required for custom provider")
        return CustomProvider(api_key=api_key, model=model, base_url=base_url)

    else:
        raise ValueError(f"Unknown provider type: {provider_type}")
