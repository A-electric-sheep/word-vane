# -*- coding: utf-8 -*-
"""文档一致性测试。

README 里写的命令参数如果和实际 argparse 定义脱节，用户照着敲就会报错。
这类"文档腐烂"很难靠人工 review 长期守住，所以用测试钉住。

这里只校验**参数名是否存在**，不校验措辞——措辞会随文案调整，参数名是契约。
"""
from __future__ import annotations

import re
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
README = REPO / 'README.md'


def _readme_text():
    return README.read_text(encoding='utf-8')


def _cli_options(func, argv=None):
    """借 --help 拿到某个 CLI 支持的选项名集合。"""
    from contextlib import redirect_stdout
    from io import StringIO

    buf = StringIO()
    with pytest.raises(SystemExit), redirect_stdout(buf):
        func(['--help'])
    text = buf.getvalue()
    return set(re.findall(r'(--[a-z][a-z0-9-]*)', text))


def test_readme_exists():
    assert README.is_file()


def test_readme_documents_real_word2md_flags():
    """README 提到的 word2md 参数必须真实存在。"""
    from word2md import main
    available = _cli_options(main)

    documented = set(re.findall(r'`(--[a-z][a-z0-9-]*)`', _readme_text()))
    # 只挑属于转换器的那些：在 README 中出现且看起来像转换器参数
    candidates = {'--verify', '--probe', '--fail-under', '--no-images',
                  '--no-toc', '--no-bold', '--code-lang', '--no-code-heuristics',
                  '--output-dir', '--version'}
    for flag in documented & candidates:
        assert flag in available, f'README 写了 {flag}，但 word2md 不支持'


def test_readme_documents_real_verify_flags():
    from verify_consistency import main
    available = _cli_options(main)
    assert '--version' in available


def test_readme_documents_real_probe_flags():
    from review_probe import main
    available = _cli_options(main)
    for flag in ('--fail-under', '--sample'):
        assert flag in available, f'probe 缺少 {flag}'


def test_readme_documents_real_batch_verify_flags():
    from batch_verify import main
    available = _cli_options(main)
    for flag in ('--fail-under', '--converted', '--workers', '--json'):
        assert flag in available, f'batch_verify 缺少 {flag}'


def test_console_scripts_declared_in_pyproject():
    """pyproject 里的入口点必须都指向真实可调用的 main。"""
    text = (REPO / 'pyproject.toml').read_text(encoding='utf-8')
    scripts = re.findall(r'^([a-z0-9-]+)\s*=\s*"([a-z_]+):(\w+)"', text, re.M)
    assert scripts, 'pyproject 未声明任何 console_scripts'

    import importlib
    for _cmd, module_name, func_name in scripts:
        mod = importlib.import_module(module_name)
        assert hasattr(mod, func_name), f'{module_name} 缺少 {func_name}'


def test_readme_has_no_unverifiable_claims():
    """README 不得出现无法复现的精确指标（如 99.998%）。

    这类数字看似权威，实则读者无从验证；一旦被追问就失信。要求改成
    附带脚本的可复现表述。
    """
    text = _readme_text()
    suspicious = re.findall(r'99\.99\d%|99\.998', text)
    assert not suspicious, f'发现不可复现的精确指标: {suspicious}'


def test_readme_states_footnote_gap():
    """脚注不支持是已知缺口，必须在限制表里如实写明。

    隐藏已知缺陷比缺陷本身更伤信任——用户踩到一次就不会再用。
    """
    text = _readme_text()
    assert '脚注' in text
    assert '不导出' in text or '不支持' in text


def test_readme_credits_docx2md_verify():
    """docx2md 早已提供 verify，README 不得暗示"唯一提供校验"。"""
    text = _readme_text()
    assert 'docx2md' in text, '应致谢并区分 docx2md 的 verify 能力'
    assert '首创' in text, '应明确说明校验非本项目首创'