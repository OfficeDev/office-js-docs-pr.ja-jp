---
title: Office JavaScript API ライブラリの参照
description: アドインで Office JavaScript API ライブラリおよび型定義を参照する方法について説明します。
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 5e26d5b0454a6833c593ff60c1577d24583dcc51
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596719"
---
# <a name="referencing-the-office-javascript-api-library"></a>Office JavaScript API ライブラリの参照

[Office JAVASCRIPT API](../reference/javascript-api-for-office.md)ライブラリには、アドインが office ホストと対話するために使用できる api が用意されています。 ライブラリを参照する最も簡単な方法は、HTML ページの`<script>` `<head>`セクション内に次のタグを追加することによって、コンテンツ配信ネットワーク (CDN) を使用する方法です。  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

これにより、アドインが最初に読み込まれたときに Office JavaScript API ファイルがダウンロードされてキャッシュされるため、指定されたバージョンの Office .js および関連付けられたファイルの最新の実装を使用していることを確認できます。

> [!IMPORTANT]
> ページのセクションの`<head>`内側から OFFICE JavaScript api を参照して、API が body 要素の前に完全に初期化されていることを確認する必要があります。 Office ホストでは、アクティブ化の 5 秒以内にアドインを初期化する必要があります。 このしきい値内にアドインがアクティブにならない場合は、応答なしが宣言され、エラー メッセージがユーザーに表示されます。

## <a name="api-versioning-and-backward-compatibility"></a>API のバージョン管理と下位互換性

前の HTML スニペットで、CDN `/1/` URL の先頭`office.js`にある、バージョン1の Office .js で最新の増分リリースを指定します。 Office JavaScript API は下位互換性を維持しているため、最新のリリースでは、以前のバージョン1で導入された API メンバーを引き続きサポートしています。 既存のプロジェクトをアップグレードする必要がある場合は、「 [Office JAVASCRIPT API およびマニフェストスキーマファイルのバージョンを更新](update-your-javascript-api-for-office-and-manifest-schema-version.md)する」を参照してください。 

AppSource から Office アドインを発行する場合は、この CDN の参照を使用する必要があります。ローカル参照は、内部シナリオ、開発シナリオ、デバッグ シナリオにのみ適用できます。

> [!NOTE]
> プレビュー API を使用するには、CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) にある Office JavaScript API ライブラリのプレビュー バージョンを参照します。

## <a name="enabling-intellisense-for-a-typescript-project"></a>TypeScript プロジェクトに対して Intellisense を有効にする

前述したように Office JavaScript API を参照するだけでなく、[指定](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js)した型定義を使用して TypeScript アドインプロジェクトの Intellisense を有効にすることもできます。 これを行うには、プロジェクトフォルダーのルートから、ノードが有効なシステムプロンプト (または git bash ウィンドウ) で次のコマンドを実行します。 (npm を含む) [Node.js](https://nodejs.org) をインストールしておく必要があります。

```command&nbsp;line
npm install --save-dev @types/office-js
```

> [!NOTE]
> プレビュー Api に対して Intellisense を有効にするには、プロジェクトフォルダーのルートで次のコマンドを実行することによって[、型定義](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js-preview)のプレビュータイプ定義を使用します。 
>
> `npm install --save-dev @types/office-js-preview`

## <a name="see-also"></a>関連項目

- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Office の JavaScript API](../reference/javascript-api-for-office.md)
