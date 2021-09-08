---
title: Office アドインのダイアログ ボックス
description: アドインのダイアログの視覚的な設計に関するベスト プラクティスOffice説明します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d674b747effa57b8a75b79f98f5ff78ccc8a92a4
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938568"
---
# <a name="dialog-boxes-in-office-add-ins"></a>Office アドインのダイアログ ボックス

ダイアログ ボックスは、作業中の Office アプリケーション ウインドウの手前に浮動するサーフェスです。ダイアログ ボックスを使用すれば、作業ウィンドウで直接開くことができないサインイン ページ、ユーザーによるアクションを確認するための要求、作業ウィンドウ内で再生すると小さすぎるビデオの表示などのタスクのために追加の画面領域を提供できます。

*図 1. ダイアログ ボックスの一般的なレイアウト*

![アプリケーションに表示されるダイアログ ボックスの一般的なOfficeです。](../images/overview-with-app-dialog.png)

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
