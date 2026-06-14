"""Smoke test for generator module."""
import json
from generator import build_system_prompt, build_user_prompt, clean_json_text, validate_report

# Test system prompt
sys_prompt = build_system_prompt()
print(f'System prompt: {len(sys_prompt)} chars')
assert 'BIRACS' in sys_prompt
assert 'h1' in sys_prompt
print('  OK: System prompt contains BIRACS + heading types')

# Test user prompt
user_prompt = build_user_prompt('测试问题：某游戏抽卡是否涉赌？')
print(f'User prompt: {len(user_prompt)} chars - OK')

# Test JSON cleaning
test_json = '```json\n{"title": "test"}\n```'
cleaned = clean_json_text(test_json)
assert cleaned == '{"title": "test"}'
print('  OK: JSON cleaning works')

# Build a good JSON report
good_json = json.dumps({
    "title": "关于线上游戏概率玩法涉赌认定的法律评估意见",
    "sections": [
        {"type": "h1", "content": "一、背景介绍"},
        {"type": "body", "content": "测试正文第一段。\n\n测试正文第二段。"},
        {"type": "h1", "content": "二、评估结论"},
        {"type": "body", "content": "结论测试正文。"},
        {"type": "h1", "content": "三、法律分析"},
        {"type": "h2", "content": "（一）测试子问题"},
        {"type": "body", "content": "分析正文。"},
        {"type": "h1", "content": "四、行业实践与司法案例"},
        {"type": "h2", "content": "（一）司法案例"},
        {"type": "table", "content": {"caption": "表1", "headers": ["案例", "法院"], "rows": [["A案", "某法院"]]}},
        {"type": "body", "content": "案例分析。"},
        {"type": "h1", "content": "五、风险评级与合规建议"},
        {"type": "h2", "content": "（一）风险评级"},
        {"type": "table", "content": {"caption": "表2", "headers": ["场景", "等级"], "rows": [["抽卡", "高风险"]]}},
        {"type": "h2", "content": "（二）合规建议"},
        {"type": "body", "content": "建议正文。"},
        {"type": "h1", "content": "六、信息来源"},
        {"type": "h2", "content": "（一）法律法规"},
        {"type": "h3", "content": "1、《刑法》第303条"},
        {"type": "body", "content": "来源：全国人大。"},
        {"type": "h2", "content": "（二）相关案例"},
        {"type": "h3", "content": "1、某案例"},
        {"type": "body", "content": "来源：中国裁判文书网。"},
        {"type": "h2", "content": "（三）其他参考文件"},
        {"type": "h3", "content": "1、某文献"},
        {"type": "body", "content": "来源：某期刊。"},
        {"type": "body", "content": "最后检索日期：2026年6月10日。"},
        {"type": "body", "content": "本报告仅供参考。"},
    ]
})
data, errors = validate_report(good_json)
print(f'Validation errors: {len(errors)}')
for e in errors:
    print(f'  - {e.field}: {e.message[:80]}')
assert len(errors) == 0, f'Expected 0 errors, got {len(errors)}'
print('  OK: Good JSON passes validation')

# Test bad JSON
bad_json = json.dumps({
    "title": "短",
    "sections": [
        {"type": "h1", "content": "**bold**"},
        {"type": "unknown_type", "content": "x"},
        {"type": "table", "content": {}},
    ]
})
data2, errors2 = validate_report(bad_json)
print(f'Bad JSON errors: {len(errors2)}')
for e in errors2:
    print(f'  - {e.field}: {e.message[:80]}')
assert len(errors2) > 0, 'Expected errors for bad JSON'
print('  OK: Bad JSON caught by validation')

# Test retry prompt
from generator import build_retry_prompt
from models import ValidationError
test_errors = [
    ValidationError(field="sections[3].content", message="包含违禁内容", section_index=3),
    ValidationError(field="sections[7].type", message="未知类型 bad_type", section_index=7),
]
retry = build_retry_prompt(test_errors, "测试问题")
assert "违禁内容" in retry
assert "bad_type" in retry
print('  OK: Retry prompt built correctly')

# Test the system prompt contains key instructions
assert "背景介绍" in sys_prompt
assert "评估结论" in sys_prompt
assert "法律分析" in sys_prompt
assert "行业实践与司法案例" in sys_prompt
assert "风险评级与合规建议" in sys_prompt
assert "信息来源" in sys_prompt
assert "footnote_on_col" in sys_prompt  # footnote column feature
assert "格式禁则" in sys_prompt
assert "裸 JSON" in sys_prompt
print('  OK: System prompt contains all 6 BIRACS sections + rules')

print()
print('=== All smoke tests passed! ===')
