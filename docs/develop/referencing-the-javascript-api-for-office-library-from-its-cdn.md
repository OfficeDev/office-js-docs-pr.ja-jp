---
title: Office JavaScript API ライブラリの参照
description: アドインで JavaScript API ライブラリOfficeタイプ定義を参照する方法について説明します。
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 5348e1acab35e01dc6c467d20a65721fb98722d47c729edeb65a2efe4a8c45f8
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57080259"
---
# <a name="referencing-the-office-javascript-api-library"></a>Office JavaScript API ライブラリの参照

[JavaScript API Officeには](../reference/javascript-api-for-office.md)、アドインがアプリケーションと対話するために使用できる API がOfficeされます。 ライブラリを参照する最も簡単な方法は、HTML ページのセクション内にCDNタグを追加してコンテンツ配信ネットワーク (CDN) `<script>` `<head>` を使用する方法です。

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

これにより、Office JavaScript API ファイルが初めて読み込まれると、Office.js の最新の実装と指定したバージョンの関連ファイルが使用されます。

> [!IMPORTANT]
> ページのセクション内Office JavaScript API を参照して、本文要素の前に API が完全に初期化 `<head>` される必要があります。

## <a name="api-versioning-and-backward-compatibility"></a>API のバージョン管理と下位互換性

前の HTML スニペットでは、URL の前に、CDN バージョン 1 の増分リリース `/1/` `office.js` を指定Office.js。 JavaScript API Office互換性が維持されるので、最新のリリースでは、バージョン 1 で以前に導入された API メンバーを引き続きサポートします。 既存のプロジェクトをアップグレードする必要がある場合は[、「JavaScript API](update-your-javascript-api-for-office-and-manifest-schema-version.md)とマニフェスト スキーマ ファイルのバージョンOffice更新する」を参照してください。 

AppSource から Office アドインを発行する場合は、この CDN の参照を使用する必要があります。ローカル参照は、内部シナリオ、開発シナリオ、デバッグ シナリオにのみ適用できます。

> [!NOTE]
> プレビュー API を使用するには、CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) にある Office JavaScript API ライブラリのプレビュー バージョンを参照します。

## <a name="enabling-intellisense-for-a-typescript-project"></a>TypeScript プロジェクトIntelliSenseを有効にする

前述のように Office JavaScript API を参照する以外に[、DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js)の型定義を使用して TypeScript アドイン プロジェクトの IntelliSense を有効にすることもできます。 これを行うには、プロジェクト フォルダーのルートからノード対応のシステム プロンプト (または git bash ウィンドウ) で次のコマンドを実行します。 (npm を含む) [Node.js](https://nodejs.org) をインストールしておく必要があります。

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>プレビュー API

新しい JavaScript API は、最初に "プレビュー" で導入され、後で十分なテストが行われるとユーザーフィードバックが必要になった後、特定の番号付き要件セットの一部になります。

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>関連項目

- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Office の JavaScript API](../reference/javascript-api-for-office.md)
