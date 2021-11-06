---
title: カスタム関数とデータ型の概要
description: カスタム関数と Office アドインで Excel データ型を使用します。
ms.date: 11/03/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 91d2fb21aae57ed7a5777136f3c4540925f339c8
ms.sourcegitcommit: 210251da940964b9eb28f1071977ea1fe80271b4
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/05/2021
ms.locfileid: "60793576"
---
# <a name="use-data-types-with-custom-functions-in-excel-preview"></a>Excel のカスタム関数でデータ型を使用する (プレビュー)

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

データ型は Excel JavaScript API を拡張して、4 種の元のデータ型 (文字列、数値、ブール値、エラー) 以外のデータ型もサポートします。 データ型には、Web イメージ、書式設定された数値、エンティティ値、エンティティ値内の配列のサポートが含まれます。

これらのデータ型はカスタム関数の能力を強化します。というのは、カスタム関数が入力値と出力値の両方としてデータ型を受け入れるからです。 カスタム関数を使用してデータ型を生成することも、既存のデータ型を関数引数として計算に取り込んだりすることもできます。 データ型の JSON スキーマが設定されると、このスキーマはカスタム関数の計算全体で維持されます。

Excel アドインでデータ型を使用する方法の詳細については、[Excel アドインのデータ型の概要](excel-data-types-overview.md)を参照してください。カスタム データ型とカスタム関数の統合の詳細については、[カスタム関数とデータ型のコア概念](custom-functions-data-types-concepts.md)を参照してください。

## <a name="see-also"></a>関連項目

* [Excel アドインのデータ型の概要](excel-data-types-overview.md)
* [Excel データ型の主要概念](excel-data-types-concepts.md)
* [カスタム関数とデータ型の主要概念](custom-functions-data-types-concepts.md)
* [Office アドインを構成して共有 JavaScript ランタイムを使用する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
