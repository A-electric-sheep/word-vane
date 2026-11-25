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


def test_pypi_name_is_not_word2md():
    """分发名不得为 `word2md`——该名字在 PyPI 已被他人占用（hnrobert/word2md）。

    这条测试防的是"有人顺手把 name 改回 word2md，发布时才发现被占用或误装
    成别人的包"。安装名用 word-vane，导入名与命令仍是 word2md。
    """
    text = (REPO / 'pyproject.toml').read_text(encoding='utf-8')
    m = re.search(r'^name\s*=\s*"([^"]+)"', text, re.M)
    assert m, 'pyproject 未声明 name'
    assert m.group(1) == 'word-vane', (
        f'分发名应为 word-vane，实际为 {m.group(1)}；'
        'PyPI 上的 word2md 属于另一个项目')


def test_readme_install_command_matches_package_name():
    """README 的 pip install 命令必须与 pyproject 的分发名一致。"""
    pkg = re.search(r'^name\s*=\s*"([^"]+)"',
                    (REPO / 'pyproject.toml').read_text(encoding='utf-8'),
                    re.M).group(1)
    text = _readme_text()
    # 注意要排除 `pip install -e ".[dev]"` 这种从当前目录安装的写法，
    # 否则会把 "." 误当成包名。
    installs = set(re.findall(
        r'pip install (?!-e )(?:"([^"]+)"|([a-zA-Z0-9_.-]+))', text))
    flat = set()
    for a, b in installs:
        token = (a or b).strip()
        if token and token != '.' and not token.startswith('.'):
            flat.add(token.split('[')[0])
    # 允许的安装目标：本包、本包带 extras、以及第三方对照项目
    third_party = {'anydoc', 'docling', 'marker-pdf', 'mineru', 'pandoc',
                   'firecrawl-anydoc', 'docx2md'}
    wrong = {i for i in flat if i not in third_party and i != pkg}
    assert not wrong, f'README 出现了非本包的安装名: {wrong}（本包应为 {pkg}）'


# ---------------------------------------------------------------- CHANGELOG
CHANGELOG = REPO / 'CHANGELOG.md'


def _changelog_text():
    return CHANGELOG.read_text(encoding='utf-8')


def test_changelog_exists():
    """CHANGELOG 是发布前的最低要求，缺失即失败。"""
    assert CHANGELOG.is_file(), '缺少 CHANGELOG.md'


def test_changelog_version_matches_source():
    """CHANGELOG 里的版本号必须与 word2md.__version__ 一致。

    这条防的是「改了版本号却忘了更新日志」——发布时最难查的一类不一致。
    """
    import word2md
    text = _changelog_text()
    # 取第一个二级标题里的版本号，如 "## [1.0.0] — 未发布"
    m = re.search(r'^##\s*\[([0-9]+\.[0-9]+\.[0-9]+)\]', text, re.M)
    assert m, 'CHANGELOG 未见形如 "## [x.y.z]" 的版本标题'
    assert m.group(1) == word2md.__version__, (
        f'CHANGELOG 版本 {m.group(1)} 与源码 {word2md.__version__} 不一致')


def test_changelog_has_required_sections():
    """变更日志至少要交代：新增、修复、限制。

    没有「已知限制」的变更日志会误导使用者以为功能完整。
    """
    text = _changelog_text()
    for section in ('### 新增', '### 修复', '### 已知限制'):
        assert section in text, f'CHANGELOG 缺少「{section}」小节'


def test_changelog_release_status_is_consistent():
    """README 与 CHANGELOG 的发布状态必须一致，且不得残留过期状态声明。

    只在「关于版本号」段与 README 首屏这类**状态声明位置**判断，
    不扫描全篇——变更条目里正常引用历史问题（如「曾写着尚未发布」）
    不应被判为违规。
    """
    changelog = _changelog_text()
    readme = _readme_text()

    released_in_cl = 'pypi.org/project/word-vane' in changelog
    released_in_readme = 'pypi.org/project/word-vane' in readme
    assert released_in_cl == released_in_readme, (
        '发布状态表述不一致：CHANGELOG=%s, README=%s'
        % (released_in_cl, released_in_readme))

    if not released_in_cl:
        return

    # CHANGELOG 的「关于版本号」段是状态声明区
    m = re.search(r'##\s*关于版本号(.*?)(?=\n##\s|\Z)', changelog, re.S)
    assert m, 'CHANGELOG 缺少「关于版本号」段'
    status_block = m.group(1)
    assert '尚未发布' not in status_block, \
        '「关于版本号」段仍声明尚未发布，与已发布状态矛盾'

    # README 首屏（版本行所在区域）不应再称开发中/未发布
    head = '\n'.join(readme.splitlines()[:8])
    assert '尚未发布' not in head, 'README 首屏仍称未发布'
    assert '开发中' not in head, 'README 首屏仍称开发中'


def test_readme_links_to_pypi():
    """README 应提供 PyPI 项目链接，便于从包页反向找到仓库。"""
    assert 'pypi.org/project/word-vane' in _readme_text(), \
        'README 缺少 PyPI 项目链接'


def test_readme_has_no_repo_relative_links():
    """README 会被渲染到 PyPI 项目页，其中不得使用仓库内相对链接。

    PyPI 不解析仓库相对路径，`](CHANGELOG.md)` 在包页上点击即 404。
    README 同时要服务于 GitHub 与 PyPI 两个场景，因此文档间跳转
    必须写成绝对 URL。
    """
    text = _readme_text()
    # 匹配 ](xxx.md) 或 ](dir/xxx.md) 这类不含协议的相对路径
    rel = re.findall(r'\]\((?!https?://|#|mailto:)([^)]*\.md)\)', text)
    assert not rel, (
        f'README 存在仓库相对链接 {rel}，在 PyPI 页面上会 404；'
        '请改为 https://github.com/A-electric-sheep/word-vane/blob/main/... ')


def test_readme_does_not_claim_unpublished():
    """已发布后，README 不得再出现「尚未发布」的过期表述。

    1.0.0 曾因先上传、后改文档，导致 PyPI 包页展示
    「开发中，尚未发布到 PyPI」与实际状态矛盾。
    """
    text = _readme_text()
    assert '尚未发布到 PyPI' not in text, \
        'README 残留「尚未发布到 PyPI」表述，与已发布状态矛盾'
    # 版本行（首屏）不应再声明「开发中」
    head = '\n'.join(text.splitlines()[:8])
    assert '开发中' not in head, 'README 首屏仍标注「开发中」，与已发布状态矛盾'


def test_changelog_documents_all_cli_entrypoints():
    """CHANGELOG 宣称的命令行入口必须与 pyproject 实际声明一致。"""
    pyproj = (REPO / 'pyproject.toml').read_text(encoding='utf-8')
    declared = set(re.findall(r'^(word2md[a-z-]*)\s*=\s*"', pyproj, re.M))
    assert declared, 'pyproject 未声明任何 word2md 入口'

    text = _changelog_text()
    missing = [c for c in declared if f'`{c}`' not in text]
    # 表格式列举也接受：只要命令名出现在文中即可
    missing = [c for c in missing if c not in text]
    assert not missing, f'CHANGELOG 未提及这些入口点: {missing}'


def test_changelog_version_heading_has_date():
    """版本标题应带发布日期（已发布状态），便于追溯。

    未发布时写「未发布」，发布后应替换为具体日期。
    """
    text = _changelog_text()
    m = re.search(r'^##\s*\[[0-9]+\.[0-9]+\.[0-9]+\]\s*[—\-–]\s*(.+)$', text, re.M)
    assert m, '版本标题缺少「— 状态/日期」部分'
    tail = m.group(1).strip()
    assert re.match(r'^\d{4}-\d{2}-\d{2}$', tail) or '未发布' in tail, (
        f'版本标题状态应为日期或「未发布」，实际: {tail}')
