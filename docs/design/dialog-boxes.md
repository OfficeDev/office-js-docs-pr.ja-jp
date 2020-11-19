---
title: Office アドインのダイアログ ボックス
description: Office アドインでのダイアログのビジュアルデザインのベストプラクティスについて説明します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: ab8ca2e768c63a53b05ed2d9ef459050455231fb
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132054"
---
# <a name="dialog-boxes-in-office-add-ins"></a>Office アドインのダイアログ ボックス

ダイアログ ボックスは、作業中の Office アプリケーション ウインドウの手前に浮動するサーフェスです。ダイアログ ボックスを使用すれば、作業ウィンドウで直接開くことができないサインイン ページ、ユーザーによるアクションを確認するための要求、作業ウィンドウ内で再生すると小さすぎるビデオの表示などのタスクのために追加の画面領域を提供できます。

*図 1. ダイアログ ボックスの一般的なレイアウト*

![Office アプリケーションに表示されるダイアログボックスの一般的なレイアウト](../images/overview-with-app-dialog.png)

## <a name="best-practices"></a>ベスト プラクティス

|するべきこと|してはいけないこと|
|:-----|:--------|
|<ul><li>アドイン名および現在のタスクを含む説明的なタイトルが含まれます。</li></ul>|<ul><li>タイトルには会社名を追加しません。</li></ul>|
||<ul><li>シナリオで必要な場合を除き、ダイアログ ボックスを開きません。</li></ul>|

## <a name="implementation"></a>実装

ダイアログ ボックスを実装するサンプルについては、GitHub の「[Office アドイン ダイアログ API の例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)」を参照してください。

## <a name="see-also"></a>関連項目

- [Dialog オブジェクト](/javascript/api/office/office.dialog)
- [Office アドインの UX 設計パターン](../design/ux-design-pattern-templates.md)
