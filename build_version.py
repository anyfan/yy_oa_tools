# UTF-8
import re
import time
import subprocess
from string import Template

# 获取当前分支最新tag string
git_tag = (
    subprocess.check_output(["git", "describe", "--tags", "--abbrev=0"])
    .decode()
    .strip()
)

git_branch = (
    subprocess.check_output(["git", "rev-parse", "--abbrev-ref", "HEAD"])
    .decode()
    .strip()
)
git_hash = (
    subprocess.check_output(["git", "rev-parse", "--short", "HEAD"])
    .decode()
    .strip()
)
git_log = (
    subprocess.check_output(["git", "log", "-10", "--pretty=format:%ad - %s", "--date=short"])
    .decode()
    .strip()
)

# 获取领先tag多少次提交 int
commits_ahead = (
    subprocess.check_output(["git", "rev-list", "HEAD", "--not", "--tags", "--count"])
    .decode()
    .strip()
)

# 获取有没有更改的文件
# Classify_list = ["α", "β", "λ", "R"]
git_status = (
    "α" if subprocess.check_output(["git", "status", "--porcelain"]).decode() else "β"
)

# git_tag = Major.Minor.Revision[classify]  v0.0.0α <<==>> 0 0 0 α
match = re.match(r"^v(\d+)\.(\d+)\.(\d+)([\D]+)$", git_tag)

version_data = {
    "git_branch": git_branch,
    "git_log": git_log,
    "git_hash": git_hash,
    "build_time": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
    "Major": int(match.group(1)),  # 主要
    "Minor": int(match.group(2)),  # 次要
    "Revision": int(match.group(3)) + int(commits_ahead),  # 修订
    "Classify": git_status,  # 分类
}
version_data["Classify_num"] = ord(version_data["Classify"])
version_data["version"] = (
    f"v{version_data['Major']}.{version_data['Minor']}.{version_data['Revision']}{version_data['Classify']}"
)

# 读取模板
with open("template/win_version.py", "rb") as fp:
    text = fp.read().decode("utf-8")

ntext = Template(text).substitute(version_data)

# 保存到文件中
with open("build/win_version.py", "w", encoding="utf-8") as fp:
    fp.write(ntext)

# 将version_data变量保存到文件中
with open("build/version_data.py", "w", encoding="utf-8") as fp:
    fp.write("version_data = " + str(version_data))
