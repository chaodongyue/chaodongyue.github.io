# Chirpy 升级步骤

本仓库使用 `chirpy-starter` 和 RubyGem 安装 Chirpy。以下步骤用于将 Chirpy 升级到新的稳定版本。

命令中的 `vX.Y.Z` 表示目标版本，例如 `v7.6.0`；执行前将所有 `vX.Y.Z` 替换成实际版本号，将 `X.Y.Z` 替换成不带 `v` 的版本号。

## 1. 检查工作区

```bash
git switch main
git status -sb
git pull --ff-only origin main
```

工作区必须干净。如果存在未提交的修改，先提交或暂存，不要带着其他修改升级。

## 2. 配置 Starter 远程

首次升级时执行：

```bash
git remote add chirpy https://github.com/cotes2020/chirpy-starter.git
git config submodule.assets/lib.ignore all
```

以后只需确认配置仍然存在：

```bash
git remote get-url chirpy
git config --get submodule.assets/lib.ignore
```

预期结果分别为：

```text
https://github.com/cotes2020/chirpy-starter.git
all
```

## 3. 查询最新稳定版本

```bash
git ls-remote --tags --refs https://github.com/cotes2020/chirpy-starter.git
```

选择版本号最大的稳定标签，不使用预发布版本。

## 4. 获取目标版本

本仓库的 `upstream` 指向 Chirpy 主题源码，可能已经存在同名标签。为了避免覆盖标签，将 Starter 标签保存为独立名称：

```bash
git fetch chirpy refs/tags/vX.Y.Z:refs/tags/chirpy-starter-vX.Y.Z
```

## 5. 创建升级分支

```bash
git switch -c upgrade-chirpy-X.Y.Z
```

## 6. 合并 Starter

```bash
git merge chirpy-starter-vX.Y.Z --squash --allow-unrelated-histories
```

出现冲突是正常情况，按下面的固定规则处理。

### assets/lib

本仓库不使用该子模块，取消暂存：

```bash
git restore --staged assets/lib
```

### 固定取新版的文件

部署工作流和依赖文件使用 Starter 新版本：

```bash
git checkout --theirs .github/workflows/pages-deploy.yml Gemfile
git add .github/workflows/pages-deploy.yml Gemfile
```

`.gitignore` 合并双方规则，不能直接覆盖已有自定义规则。

### 固定保留本仓库的文件

About 页面和 README 包含本仓库自己的内容：

```bash
git checkout --ours _tabs/about.md README.md
git add _tabs/about.md README.md
```

### _config.yml

保留以下本地配置：

- `lang`、`timezone`
- `title`、`tagline`、`description`、`url`
- GitHub 用户名和社交信息
- 搜索引擎验证码
- 头像
- Giscus 配置
- `_drafts` 默认配置

同时从 Starter 新版本加入新增的配置项。处理完成后执行：

```bash
git add _config.yml .gitignore
```

确认没有未解决冲突：

```bash
git status
git diff --check
```

`git status` 中不能再出现 `Unmerged paths`。

## 7. 更新 Ruby 依赖

```bash
bundle update
bundle exec ruby -e 'puts Gem::Specification.find_by_name("jekyll-theme-chirpy").version'
```

输出版本应为目标版本 `X.Y.Z`。

本仓库使用 Starter/RubyGem 方式，不需要执行 `npm run build`，也不需要提交 `_sass/vendors` 或 `assets/js/dist`。

## 8. 测试站点

```bash
bash tools/test.sh
```

必须确保 Jekyll 构建和 HTML-Proofer 全部通过。

## 9. 提交升级

```bash
git add .
git commit -m "chore: upgrade to X.Y.Z"
```

检查升级提交：

```bash
git status -sb
git diff main...HEAD --stat
```

## 10. 合并到 main

```bash
git switch main
git merge --ff-only upgrade-chirpy-X.Y.Z
```

## 11. 推送

```bash
git push origin main
```

推送完成后检查 GitHub Pages 工作流是否成功。

## 升级失败时恢复

如果 squash merge 产生的冲突不准备继续处理，且没有需要保留的本地修改，可执行：

```bash
git reset --merge HEAD
git switch main
```

不要使用 `git reset --hard`，避免误删本地修改。
