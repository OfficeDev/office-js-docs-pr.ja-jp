---
title: カスタム関数とデータ型のコア概念
description: カスタム関数でデータ型を使用するExcel概念について説明します。
ms.date: 11/03/2021
ms.topic: conceptual
ms.custom: scenarios:getting-started
ms.localizationpriority: medium
ms.openlocfilehash: 3b7e735f78ca7b6dcdffa3bd5e8ba9c9d3093766
ms.sourcegitcommit: ad5d7ab21f64012543fb2bd9226d90330d25468b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/04/2021
ms.locfileid: "60749407"
---
# <a name="custom-functions-and-data-types-core-concepts-preview"></a>カスタム関数とデータ型のコア概念 (プレビュー)

[!include[Custom functions and data types availability note](../includes/excel-custom-functions-data-types-note.md)]

データ型は、Excel 4 (文字列、数値、ブール値、エラー) を超えるデータ型のサポートを拡張することで、JavaScript API の機能を強化します。 データ型には、書式設定された数値、Web イメージ、エンティティ値、エンティティ値内の配列のサポートが含まれます。 カスタム関数は、入力値と出力値の両方としてデータ型を受け入れ、カスタム関数の計算力を拡張します。

アドインでデータ型を使用する方法Excel詳細については、「Excelデータ型のコア[概念」を参照してください](excel-data-types-concepts.md)。

## <a name="how-custom-functions-handle-data-types"></a>カスタム関数がデータ型を処理する方法

カスタム関数は、データ型を認識し、パラメーター値として受け入れることができます。 カスタム関数は、戻り値の新しいデータ型を作成できます。 カスタム関数は、JavaScript API のデータ型と同じ JSON スキーマをExcel、この JSON スキーマはカスタム関数の計算および評価として維持されます。

> [!NOTE]
> カスタム関数は、データ型によって提供される拡張エラー オブジェクトの完全な機能をサポートしません。 カスタム関数は、データ型エラー オブジェクトを受け入れできますが、計算中は保持されません。 現時点では、カスタム関数は [CustomFunctions.Error](custom-functions-errors.md)オブジェクトに含まれるエラーのみをサポートします。

## <a name="enable-data-types-for-custom-functions"></a>カスタム関数のデータ型を有効にする

この機能を使用するには、JSON メタデータを手動で更新する必要があります。 一時的なテストを行う場合は、JSON メタデータを手動でScript Labの設定をカスタマイズできます。 以下のセクションでは、これらの手順の概要を詳しく説明します。

### <a name="manually-update-json-metadata"></a>JSON メタデータを手動で更新する

カスタム関数プロジェクトには、JSON メタデータ ファイルが含まれます。 この JSON メタデータ ファイルは、データ型 API で使用される JSON スキーマとは異なります。 カスタム関数とデータ型の統合を使用するには、カスタム関数 JSON メタデータ ファイルを手動で更新してプロパティを含める必要があります `allowCustomDataForDataTypeAny` 。 このプロパティをに設定します `true` 。

手動 JSON 作成プロセスの詳細については、「カスタム関数の JSON メタデータを手動 [で作成する」を参照してください](custom-functions-json.md)。 この [プロパティの詳細については、「allowCustomDataForDataTypeAny」](custom-functions-json.md#allowcustomdatafordatatypeany-preview) を参照してください。

### <a name="script-lab-option"></a>Script Labオプション

データ型とのカスタム関数の統合は、前のセクションで説明した手動の JSON メタデータ更新Script Labに加えて、データ型とのテストに使用できます。 この方法の詳細については、「Script Lab JavaScript API を使用して JavaScript API をOffice[する」を参照Script Lab。](../overview/explore-with-script-lab.md) この機能をテストするには、Script Labを使用して設定を更新します。

1. [コード] 作業Script Lab **を** 開きます。
1. 右下隅にある [ウィンドウ] ボタン **設定** します。
1. [ユーザー]**タブに移動設定** 入力します `allowCustomDataForDataTypeAny: true` 。

![カスタム関数のデータ型を有効にする手順を示すスクリーンショットScript Lab。](../images/custom-functions-script-lab-data-type.png)

## <a name="output-a-formatted-number-value"></a>書式設定された数値を出力する

次のコード サンプルは、カスタム関数を使用して [FormattedNumberCellValue](/javascript/api/excel/excel.formattednumbercellvalue) データ型を作成する方法を示しています。 この関数は、基本数値と書式設定を入力パラメーターとして受け取り、書式設定された数値データ型を出力として返します。

```js
/**
 * Take a number as the input value and return a formatted number value as the output.
 * @customfunction
 * @param {number} value
 * @param {string} format (e.g. "0.00%")
 * @returns A formatted number value.
 */
function createFormattedNumber(value, format) {
    return {
        type: "FormattedNumber",
        basicValue: value,
        numberFormat: format
    }
}
```

## <a name="input-an-entity-value"></a>エンティティ値の入力

次のコード サンプルは [、EntityCellValue](/javascript/api/excel/excel.entitycellvalue) データ型を入力として受け取るカスタム関数を示しています。 パラメーターが `attribute` に設定されている `text` 場合、関数はエンティティ値 `text` のプロパティを返します。 それ以外の場合、関数はエンティティ `basicValue` 値のプロパティを返します。

```js
/**
 * Accept an entity value data type as a function input.
 * @customfunction
 * @param {any} value
 * @param {string} attribute
 * @returns {any} The text value of the entity.
 */
function getEntityAttribute(value, attribute) {
    if (value.type == "Entity") {
        if (attribute == "text") {
            return value.text;
        } else {
            return value.properties[attribute].basicValue;
        }
    } else {
        return JSON.stringify(value);
    }
}
```

## <a name="see-also"></a>関連項目

* [カスタム関数とデータ型の概要](custom-functions-data-types-overview.md)
* [Excel アドインのデータ型の概要](excel-data-types-overview.md)
* [Excelデータ型のコア概念](excel-data-types-concepts.md)
* [Office アドインを構成して共有 JavaScript ランタイムを使用する](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
