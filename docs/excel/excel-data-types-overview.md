---
title: Excel アドインのデータ型の概要
description: Excel JavaScript API のデータ型を使用すると、Office アドイン開発者は、書式設定された数値、Web イメージ、エンティティ値、エンティティ値内の配列、および拡張エラーをデータ型として操作できます。
ms.date: 11/01/2021
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: f5866b3ec27fc2e5869150feb45564701824afcd
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681826"
---
# <a name="overview-of-data-types-in-excel-add-ins-preview"></a>Excel アドインのデータ型の概要 (プレビュー)

> [!NOTE]
> 現在、データ型 API はパブリック プレビューでのみ使用できます。 プレビュー API は変更されることがあります。運用環境での使用は意図されていません。 運用環境やビジネス上重要なドキュメントでプレビュー API を使用しないでください。

> [!IMPORTANT]
> `Range.valuesAsJSON` などの一部のデータ型 API は、アクティブな開発中であり、パブリック プレビューではまだ利用できません。 この記事は、概念的な紹介を目的としています。 この記事で説明されている、パブリック プレビューにはまだ含まれていない概念は、間もなくプレビューにリリースされる予定です。

Excel JavaScript API のデータ型を使用すると、アドイン開発者は、書式設定された数値、Web イメージ、エンティティ値などのオブジェクトとして複雑なデータ構造を整理できます。

データ型を追加する前は、Excel JavaScript API でサポートされていたのは、文字列、数値、ブール値、エラーデータ型でした。 Excel UI 書式設定レイヤーでは、元からある 4 種のデータ型を含むセルに通貨、日付、およびその他の種類の書式を追加できますが、この書式設定レイヤーは Excel UI 上の元のデータ型の表示のみを制御します。 Excel UI のセルが通貨または日付として書式設定されている場合でも、基になる数値は変更されません。 基になる値と Excel UI の書式設定された表示の間のこのギャップにより、アドインの計算中に混乱やエラーが発生する可能性があります。 このギャップの解決策としては、カスタム データ型を使用することです。

データ型は、4 種の元のデータ型 (文字列、数値、ブール値、エラー) を超えて Excel JavaScript API のサポートを拡張し、Web イメージ、書式設定された数値、エンティティ値、エンティティ値内の配列、および強化されたエラー データ型を柔軟なデータ構造として含めます。 これらの型は、多くの [linked data types](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772) エクスペリエンスを強化し、アドインの計算中の精度と簡易性を実現し、Excel アドインの可能性を 2 次元グリッドを超えて拡張します。

## <a name="data-types-and-custom-functions"></a>データ型とカスタム関数

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

データ型は、カスタム関数の機能を強化します。 カスタム関数は、カスタム関数への入力とカスタム関数の出力の両方としてデータ型を受け取り、カスタム関数は Excel JavaScript API と同じ JSON スキーマをデータ型に使用します。 このデータ型の JSON スキーマは、カスタム関数により計算および評価がされるときに維持されます。 データ型とカスタム関数の統合の詳細については、「[カスタム関数とデータ型の主要概念](/custom-functions-data-types-concepts.md)」を参照してください。

## <a name="see-also"></a>関連項目

* [Excel データ型の主要概念](/excel-data-types-concepts.md)
* [Excel JavaScript API リファレンス](../reference/overview/excel-add-ins-reference-overview.md)
* [カスタム関数とデータ型の概要](/custom-functions-data-types-overview.md)