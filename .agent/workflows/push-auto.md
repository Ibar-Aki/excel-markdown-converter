---
description: 変更をコミットしてGitHubへプッシュする自動化ワークフロー
---

# /push-auto

現在のvoid-curiosityリポジトリの変更をすべてコミットし、リモートリポジトリ（GitHub）のmainブランチへ自動的にプッシュします。
このワークフローは、コマンドの確認プロンプトをスキップしてすべて自動実行（turbo-all）されます。

// turbo-all

1. 現在の変更状態を確認します。

```powershell
git status
```

1. すべての変更をステージングします。

```powershell
git add .
```

1. 変更をコミットします。

```powershell
git commit -m "Auto-commit by Antigravity /push-auto workflow"
```

`

> [!NOTE]
> コミットメッセージを指定したい場合は、通常の `run_command` を使用するか、このファイルを手動で編集してください。
