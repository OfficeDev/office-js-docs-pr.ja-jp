---
title: Excel アドインのデータ型の概要
description: Excel JavaScript API のデータ型を使用すると、Office アドイン開発者は、書式設定された数値、Web イメージ、エンティティ値、エンティティ値内の配列、および拡張エラーをデータ型として操作できます。
ms.date: 12/27/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 06a10051b1b243689f9d46d22c38cbdbfb155e4d
ms.sourcegitcommit: b46d2afc92409bfc6612b016b1cdc6976353b19e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/30/2021
ms.locfileid: "61647952"
---
# <a name="overview-of-data-types-in-excel-add-ins-preview"></a>Excel アドインのデータ型の概要 (プレビュー)

> [!NOTE]
> 現在、データ型 API はパブリック プレビューでのみ使用できます。 プレビュー API は変更されることがあります。運用環境での使用は意図されていません。 試用はテスト環境と開発環境に限定することをお勧めします。 運用環境やビジネス上重要なドキュメントでプレビュー API を使用しないでください。
>
> プレビュー API を使用するには:
>
> - CDN (**の** ベータhttps://appsforoffice.microsoft.com/lib/beta/hosted/office.js) ライブラリを参照する必要があります。 TypeScript コンパイルおよび IntelliSense の [型定義ファイル](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)は CDN で見つかり、[DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) にあります。 これらの型は、`npm install --save-dev @types/office-js-preview` を使用してインストールできます。 詳細については、[@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) NPM パッケージ readme を参照してください。
> - 最新の Office ビルドにアクセスするには、[Office Insider プログラム](https://insider.office.com)に参加する必要がある場合もあります。
>
> Windows 版 Office でデータ型を試すには、16.0.14626.10000 以上の Excel ビルド番号が必要です。 Office on Mac でデータ型を試すには、16.55.21102600 以上の Excel ビルド番号が必要です。

Excel JavaScript API のデータ型を使用すると、アドイン開発者は、書式設定された数値、Web イメージ、エンティティ値などのオブジェクトとして複雑なデータ構造を整理できます。

データ型を追加する前は、Excel JavaScript API でサポートされていたのは、文字列、数値、ブール値、エラーデータ型でした。 Excel UI 書式設定レイヤーでは、元からある 4 種のデータ型を含むセルに通貨、日付、およびその他の種類の書式を追加できますが、この書式設定レイヤーは Excel UI 上の元のデータ型の表示のみを制御します。 Excel UI のセルが通貨または日付として書式設定されている場合でも、基になる数値は変更されません。 基になる値と Excel UI の書式設定された表示の間のこのギャップにより、アドインの計算中に混乱やエラーが発生する可能性があります。 このギャップの解決策としては、カスタム データ型を使用することです。

データ型は、4 種の元のデータ型 (文字列、数値、ブール値、エラー) を超えて Excel JavaScript API のサポートを拡張し、Web イメージ、書式設定された数値、エンティティ値、エンティティ値内の配列、および強化されたエラー データ型を柔軟なデータ構造として含めます。 これらの型は、多くの [linked data types](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772) エクスペリエンスを強化し、アドインの計算中の精度と簡易性を実現し、Excel アドインの可能性を 2 次元グリッドを超えて拡張します。

## <a name="data-types-and-custom-functions"></a>データ型とカスタム関数

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

データ型は、カスタム関数の機能を強化します。 カスタム関数は、カスタム関数への入力とカスタム関数の出力の両方としてデータ型を受け取り、カスタム関数は Excel JavaScript API と同じ JSON スキーマをデータ型に使用します。 このデータ型の JSON スキーマは、カスタム関数により計算および評価がされるときに維持されます。 データ型とカスタム関数の統合の詳細については、「[カスタム関数とデータ型](custom-functions-data-types-concepts.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Excel データ型の主要概念](excel-data-types-concepts.md)
- [Excel JavaScript API リファレンス](../reference/overview/excel-add-ins-reference-overview.md)
- [カスタム関数とデータ型](custom-functions-data-types-concepts.md)