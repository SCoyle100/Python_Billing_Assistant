import os
import unittest
from unittest.mock import MagicMock, patch

from utils.openai_json import chat_completion_json


class TestOpenAIJson(unittest.TestCase):
    @patch("utils.openai_json.OpenAI")
    def test_uses_terra_with_non_reasoning_chat_completions(self, mock_openai):
        response = MagicMock()
        response.choices[0].message.content = '{"invoices": []}'
        mock_openai.return_value.chat.completions.create.return_value = response

        with patch.dict(os.environ, {}, clear=True):
            result = chat_completion_json(
                "Extract invoices.",
                "Invoice text",
                max_completion_tokens=321,
            )

        self.assertEqual(result, {"invoices": []})
        request = mock_openai.return_value.chat.completions.create.call_args.kwargs
        self.assertEqual(request["model"], "gpt-5.6-terra")
        self.assertEqual(request["reasoning_effort"], "none")
        self.assertEqual(request["max_completion_tokens"], 321)
        self.assertNotIn("max_tokens", request)

    @patch("utils.openai_json.OpenAI")
    def test_older_environment_override_does_not_receive_56_reasoning_setting(self, mock_openai):
        response = MagicMock()
        response.choices[0].message.content = "{}"
        mock_openai.return_value.chat.completions.create.return_value = response

        with patch.dict(os.environ, {"OPENAI_CHAT_MODEL": "gpt-4.1-mini"}, clear=True):
            chat_completion_json("Extract invoices.", "Invoice text")

        request = mock_openai.return_value.chat.completions.create.call_args.kwargs
        self.assertEqual(request["model"], "gpt-4.1-mini")
        self.assertNotIn("reasoning_effort", request)


if __name__ == "__main__":
    unittest.main()
