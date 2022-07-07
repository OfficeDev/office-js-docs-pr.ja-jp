---
title: Office JavaScript API ライブラリの参照
description: アドインで Office JavaScript API ライブラリと型定義を参照する方法について説明します。
ms.date: 02/18/2021
ms.localizationpriority: medium
ms.openlocfilehash: 38121fe3d3df0a86fef3e2c8e3a58399640f1e2a
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660117"
---
# <a name="referencing-the-office-javascript-api-library"></a>Office JavaScript API ライブラリの参照

[Office JavaScript API](../reference/javascript-api-for-office.md) ライブラリには、アドインが Office アプリケーションとの対話に使用できる API が用意されています。 ライブラリを参照する最も簡単な方法は、HTML ページのセクションに次 `<script>` のタグを追加して、コンテンツ配信ネットワーク (CDN) を `<head>` 使用することです。

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

これにより、アドインが初めて読み込まれると、Office JavaScript API ファイルがダウンロードされてキャッシュされ、指定されたバージョンのOffice.jsとその関連ファイルの最新の実装が使用されていることを確認できます。

> [!IMPORTANT]
> 本文要素の前に API が完全に `<head>` 初期化されるようにするには、ページのセクション内から Office JavaScript API を参照する必要があります。

## <a name="api-versioning-and-backward-compatibility"></a>API のバージョン管理と下位互換性

前の HTML スニペットでは、 `/1/` CDN URL の `office.js` 前に、Office.jsのバージョン 1 内の最新の増分リリースが指定されています。 Office JavaScript API は下位互換性を維持するため、最新リリースでは、バージョン 1 の前に導入された API メンバーが引き続きサポートされます。 既存のプロジェクトをアップグレードする必要がある場合は、「 [Office JavaScript API とマニフェスト スキーマ ファイルのバージョンを更新](update-your-javascript-api-for-office-and-manifest-schema-version.md)する」を参照してください。 

AppSource から Office アドインを発行する場合は、この CDN の参照を使用する必要があります。ローカル参照は、内部シナリオ、開発シナリオ、デバッグ シナリオにのみ適用できます。

> [!NOTE]
> プレビュー API を使用するには、CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) にある Office JavaScript API ライブラリのプレビュー バージョンを参照します。

## <a name="enabling-intellisense-for-a-typescript-project"></a>TypeScript プロジェクトで IntelliSense を有効にする

前述のように Office JavaScript API を参照するだけでなく、 [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js) の型定義を使用して、IntelliSense for TypeScript アドイン プロジェクトを有効にすることもできます。 これを行うには、プロジェクト フォルダーのルートからノード対応のシステム プロンプト (または git bash ウィンドウ) で次のコマンドを実行します。 (npm を含む) [Node.js](https://nodejs.org) をインストールしておく必要があります。

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>プレビュー API

新しい JavaScript API は、最初に "プレビュー" で導入され、後で、十分なテストが行われ、ユーザーフィードバックが取得された後に、特定の番号付き要件セットの一部になります。

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>関連項目

- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Office の JavaScript API](../reference/javascript-api-for-office.md)
