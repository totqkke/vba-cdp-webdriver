# vba-cdp-webdriver

Chromium ベースのブラウザを VBA から操作するための派生版プロジェクトです。  
元プロジェクトをベースに、公開しやすい形へ整理した VBA ソース一式をまとめています。

## Base project
このプロジェクトは、以下を出発点とした派生版です。

- 24000/ChromeControler-No-Selenium-WebDriver-VBAJSON

元プロジェクトをベースにしつつ、今回の公開対象に合わせて構成整理と調整を行っています。

## Overview
このリポジトリは、CDP を利用したブラウザ自動操作を VBA で扱うための派生版です。  
Selenium / WebDriver の追加導入が難しい環境でも、VBA ベースでブラウザ操作を組み立てたい場合の出発点として使うことを想定しています。

## What this project is
このプロジェクトでは、Chromium ベースのブラウザを VBA から操作するためのクラス、モジュール、補助コードを公開しています。  
派生版として、公開向けに整理した構成で参照できるようにしています。

## What is included in this derived version
- ブラウザ操作のための主要な VBA ソース
- 派生版として整理したクラス / モジュール構成
- サンプルコードを含む補助モジュール群
- 公開リポジトリとして見通しやすい最小構成


## Repository structure
- `README.md`  
  このリポジトリの概要です。

- `LICENSE`  
  ライセンス情報です。

- `src/`  
  主要な VBA ソースです。  
  クラスや中核となる処理を確認する場合は、まずここを見てください。

- `Module/`  
  補助モジュール群です。  
  `Sample` もここに含まれており、実際の使い方や呼び出し方の入口として確認できます。

## Where to start
最初に全体像をつかむ場合は、次の順番で見る想定です。

1. `README.md` で概要を確認する  
2. `src/` で主要クラスと中核処理を見る  
3. `Module/` の `Sample` を見て、実際の呼び出し方を確認する  

## Notes
この README では、まずリポジトリ全体の位置づけと入口が分かることを優先しています。  
実装の詳細、変更点の技術的な説明、設計上の意図は別記事側に切り分ける想定です。

## License
This repository is published under the MIT License.  
See `LICENSE`.
