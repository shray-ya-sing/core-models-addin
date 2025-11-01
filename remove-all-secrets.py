#!/usr/bin/env python3
import sys
import re

# Read the file content from stdin
content = sys.stdin.read()

# Remove Anthropic API keys
content = re.sub(
    r"anthropicApiKey: process\.env\.ANTHROPIC_API_KEY \|\| 'sk-ant-[^']+',",
    "anthropicApiKey: process.env.ANTHROPIC_API_KEY || '',",
    content
)

# Remove OpenAI API keys (multiple patterns)
content = re.sub(
    r"openaiApiKey: process\.env\.OPENAI_API_KEY \|\| 'sk-proj-[^']+',",
    "openaiApiKey: process.env.OPENAI_API_KEY || '',",
    content
)

content = re.sub(
    r"OPENAI_API_KEY\s*=\s*'sk-proj-[^']+",
    "OPENAI_API_KEY=os.environ.get('OPENAI_API_KEY', '')",
    content
)

content = re.sub(
    r"const apiKey = 'sk-proj-[^']+';",
    "const apiKey = process.env.OPENAI_API_KEY || '';",
    content
)

# Write the modified content to stdout
sys.stdout.write(content)
