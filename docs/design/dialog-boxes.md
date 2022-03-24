---
title: Office アドインのダイアログ ボックス
description: アドインのダイアログを視覚的に設計するためのベスト プラクティスOffice説明します。
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 623a94b5cd0fdd398de2e61cb779f5d256c46d76
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743255"
---
# <a name="dialog-boxes-in-office-add-ins"></a>Office アドインのダイアログ ボックス

ダイアログ ボックスは、作業中の Office アプリケーション ウインドウの手前に浮動するサーフェスです。ダイアログ ボックスを使用すれば、作業ウィンドウで直接開くことができないサインイン ページ、ユーザーによるアクションを確認するための要求、作業ウィンドウ内で再生すると小さすぎるビデオの表示などのタスクのために追加の画面領域を提供できます。

*図 1. ダイアログ ボックスの一般的なレイアウト*

![アプリケーションに表示されるダイアログ ボックスの一般的Officeレイアウト。](../images/overview-with-app-dialog.png)

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
