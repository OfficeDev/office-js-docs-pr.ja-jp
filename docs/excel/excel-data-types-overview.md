---
title: Excel アドインのデータ型の概要
description: Excel JavaScript API のデータ型を使用すると、Office アドイン開発者は、書式設定された数値、Web イメージ、エンティティ、エンティティ内の配列、および拡張エラーをデータ型として操作できます。
ms.date: 10/10/2022
ms.topic: conceptual
ms.prod: excel
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 2d19eacc23d64f472f32363fc93155b6e023ba04
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2022
ms.locfileid: "68540978"
---
# <a name="overview-of-data-types-in-excel-add-ins"></a>Excel アドインのデータ型の概要

データ型は、複雑なデータ構造をオブジェクトとして整理します。 これには、書式設定された数値、Web イメージ、エンティティ カードとしてのエンティティが含 [まれます](excel-data-types-entity-card.md)。

データ型を追加する前は、Excel JavaScript API でサポートされていたのは、文字列、数値、ブール値、エラーデータ型でした。 Excel UI 書式設定レイヤーでは、元からある 4 種のデータ型を含むセルに通貨、日付、およびその他の種類の書式を追加できますが、この書式設定レイヤーは Excel UI 上の元のデータ型の表示のみを制御します。 Excel UI のセルが通貨または日付として書式設定されている場合でも、基になる数値は変更されません。 基になる値と Excel UI の書式設定された表示の間のこのギャップにより、アドインの計算中に混乱やエラーが発生する可能性があります。 データ型 API は、このギャップの解決策です。

データ型は、4 つの元のデータ型 (文字列、数値、ブール値、エラー) を超えて Excel JavaScript API のサポートを拡張し、 [Web イメージ](excel-data-types-concepts.md#web-image-values)、 [書式設定された数値](excel-data-types-concepts.md#formatted-number-values)、 [エンティティ、エンティティ](excel-data-types-concepts.md#entity-values)内の配列、柔軟なデータ構造としての [エラー データ型](excel-data-types-concepts.md#improved-error-support) の改善を含めます。 これらの型は、多くの [linked data types](https://support.microsoft.com/office/what-linked-data-types-are-available-in-excel-6510ab58-52f6-4368-ba0f-6a76c0190772) エクスペリエンスを強化し、アドインの計算中の精度と簡易性を実現し、Excel アドインの可能性を 2 次元グリッドを超えて拡張します。

データ型 API を使用する方法については、 [Excel データ型のコア概念](excel-data-types-concepts.md) に関する記事を参照してください。

## <a name="data-types-and-custom-functions"></a>データ型とカスタム関数

データ型は、カスタム関数の機能を強化します。 カスタム関数は、カスタム関数への入力とカスタム関数の出力の両方としてデータ型を受け取り、カスタム関数は Excel JavaScript API と同じ JSON スキーマをデータ型に使用します。 このデータ型の JSON スキーマは、カスタム関数により計算および評価がされるときに維持されます。 データ型とカスタム関数の統合の詳細については、「[カスタム関数とデータ型](custom-functions-data-types-concepts.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Excel データ型の主要概念](excel-data-types-concepts.md)
- [エンティティ値データ型でカードを使用する](excel-data-types-entity-card.md)
- [Excel JavaScript API リファレンス](../reference/overview/excel-add-ins-reference-overview.md)
- [カスタム関数とデータ型](custom-functions-data-types-concepts.md)