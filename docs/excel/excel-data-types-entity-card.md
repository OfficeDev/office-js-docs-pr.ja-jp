---
title: Excel JavaScript API データ型エンティティ値カード
description: Excel アドインのデータ型でエンティティ 値カードを使用する方法について説明します。
ms.date: 05/19/2022
ms.topic: conceptual
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7f9b2c146826c8247abee6ece105d04a335c41f1
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628168"
---
# <a name="use-cards-with-entity-value-data-types-preview"></a>エンティティ値データ型でカードを使用する (プレビュー)

[!include[Data types preview availability note](../includes/excel-data-types-preview.md)]

この記事では、[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) を使用して、エンティティ値データ型を持つExcel UI でカード モーダル ウィンドウを作成する方法について説明します。 これらのカードは、関連する画像、製品カテゴリ情報、データ属性など、セルに既に表示されている情報を超えて、エンティティ値に含まれる追加情報を表示できます。

エンティティ値 ( [EntityCellValue](/javascript/api/excel/excel.entitycellvalue)) は、データ型のコンテナーであり、オブジェクト指向プログラミングのオブジェクトに似ています。 この記事では、エンティティ値カードのプロパティ、レイアウト オプション、およびデータ属性機能を使用して、カードとして表示されるエンティティ値を作成する方法について説明します。

次のスクリーンショットは、オープン エンティティ値カードの例を示しています。この例では、食品店の製品の一覧の **Tofu** 製品です。

:::image type="content" source="../images/excel-data-types-entity-card-tofu.png" alt-text="カード ウィンドウが表示されたエンティティ値データ型を示すスクリーンショット。":::

## <a name="card-properties"></a>カードのプロパティ

entity value [`properties`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-properties-member) プロパティを使用すると、データ型に関するカスタマイズされた情報を設定できます。 キーは `properties` 入れ子になったデータ型を受け入れます。 入れ子になった各プロパティ (データ型) には、設定と`basicValue`設定が必要です`type`。

> [!IMPORTANT]
> 入れ子になった `properties` データ型は、後続の記事セクションで説明する [カード レイアウト](#card-layout) 値と組み合わせて使用されます。 入れ子になったデータ型を定義した後、カードに `properties`表示するには、プロパティに `layouts` データ型を割り当てる必要があります。

次のコード スニペットは、複数のデータ型が入れ子になっているエンティティ値の JSON を示 `properties`しています。

> [!NOTE]
> 完全なコード サンプルでこの JSON を使用する方法については、 [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/excel/85-preview-apis/data-types-entity-values.yaml) リポジトリを参照してください。

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        "Product ID": {
            type: Excel.CellValueType.string,
            basicValue: productID.toString() || ""
        },
        "Product Name": {
            type: Excel.CellValueType.string,
            basicValue: productName || ""
        },
        "Quantity Per Unit": {
            type: Excel.CellValueType.string,
            basicValue: product.quantityPerUnit || ""
        },
        "Unit Price": {
            type: Excel.CellValueType.formattedNumber,
            basicValue: product.unitPrice,
            numberFormat: "$* #,##0.00"
        },
        Discontinued: {
            type: Excel.CellValueType.boolean,
            basicValue: product.discontinued || false
        }
    },
    layouts: {
        // Enter layout settings here.
    }
};
```

次のスクリーンショットは、上記のコード スニペットを使用するエンティティ値カードを示しています。 スクリーンショットは、前のコード スニペットの **製品 ID**、 **製品名**、 **ユニットあたりの数量**、 **単価** の情報を示しています。

:::image type="content" source="../images/excel-data-types-entity-card-properties.png" alt-text="カード レイアウト ウィンドウが表示されたエンティティ値データ型を示すスクリーンショット。カードには、製品名、製品 ID、ユニットあたりの数量、単価の情報が表示されます。":::

## <a name="card-layout"></a>カードレイアウト

エンティティ値 [`layouts`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-layouts-member) プロパティは、エンティティに対して a を [`card`](/javascript/api/excel/excel.entityviewlayouts) 作成し、カードのタイトル、カードの画像、表示するセクションの数など、そのカードの外観を指定します。

> [!IMPORTANT]
> 入れ子になった `layouts` 値は、前の記事セクションで説明した [Card プロパティ](#card-properties) データ型と組み合わせて使用されます。 入れ子になったデータ型は、カードに `properties` 表示するために割り当てる `layouts` 前に定義する必要があります。

プロパティ内で `card` 、オブジェクトを [`CardLayoutStandardProperties`](/javascript/api/excel/excel.cardlayoutstandardproperties) 使用して、カードのコンポーネント (例: `title`, `subTitle`、 `sections`.

次のエンティティ値 JSON コード スニペットは、 `card` 入れ子になった `title` オブジェクトとカード内の 3 つの `sections` レイアウトを示しています。 `title`このプロパティには、前の「[Card プロパティ](#card-properties)`"Product Name"`」の記事セクションで対応するデータ型があることに注意してください。 プロパティは `sections` 入れ子になった配列を受け取り、オブジェクトを [`CardLayoutSectionStandardProperties`](/javascript/api/excel/excel.cardlayoutsectionstandardproperties) 使用して各セクションの外観を定義します。

各カード セクション内で、次`title`のような`layout`要素を`properties`指定できます。 キーは `layout` オブジェクトを [`CardLayoutListSection`](/javascript/api/excel/excel.cardlayoutlistsection) 使用し、値 `"List"`を受け入れます。 キーは `properties` 文字列の配列を受け入れます。 などの`"Product ID"`値には、前の[「カードプロパティ](#card-properties)」の記事セクションで対応するデータ型があることに`properties`注意してください。 セクションは折りたたみ可能で、エンティティ カードを Excel UI で開いたときに、ブール値を折りたたんだり折りたたんだりしないように定義することもできます。

> [!NOTE]
> 完全なコード サンプルでこの JSON を使用する方法については、 [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/excel/85-preview-apis/data-types-entity-values.yaml) リポジトリを参照してください。

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        card: {
            title: { 
                property: "Product Name" 
            },
            sections: [
                {
                    layout: "List",
                    properties: ["Product ID"]
                },
                {
                    layout: "List",
                    title: "Quantity and price",
                    collapsible: true,
                    collapsed: false, // This section will not be collapsed when the card is opened.
                    properties: ["Quantity Per Unit", "Unit Price"]
                },
                {
                    layout: "List",
                    title: "Additional information",
                    collapsible: true,
                    collapsed: true, // This section will be collapsed when the card is opened.
                    properties: ["Discontinued"]
                }
            ]
        }
    }
};
```

次のスクリーンショットは、上記のコード スニペットを使用するエンティティ値カードを示しています。 スクリーンショットは、**Product Name** を`title`使用し、**Pavlova** に設定されているオブジェクトを示しています。 スクリーンショットには 、 `sections`. [ **数量と価格]** セクションは折りたたみ可能で、 **ユニットあたりの数量** と **単価が** 含まれています。 **[追加情報]** フィールドは折りたたみ可能で、カードを開いたときに折りたたまれます。

:::image type="content" source="../images/excel-data-types-entity-card-sections.png" alt-text="カード レイアウト ウィンドウが表示されたエンティティ値データ型を示すスクリーンショット。カードには、カードのタイトルとセクションが表示されます。":::

## <a name="card-data-attribution"></a>カード データ属性

エンティティ値カードは、データ属性を表示して、エンティティ カード内の情報のプロバイダーにクレジットを提供できます。 エンティティ値[`provider`](/javascript/api/excel/excel.entitycellvalue#excel-excel-entitycellvalue-provider-member)プロパティでは、オブジェクトを[`CellValueProviderAttributes`](/javascript/api/excel/excel.cellvalueproviderattributes)使用します。このオブジェクトは、そのオブジェクト `logoSourceAddress`、および値を`description``logoTargetAddress`定義します。

データ プロバイダー プロパティは、エンティティ カードの左下隅に画像を表示します。 これを使用して、 `logoSourceAddress` イメージのソース URL を指定します。 ロゴ イメージが選択されている場合、この値は `logoTargetAddress` URL 変換先を定義します。 `description`ロゴの上にマウス ポインターを置くと、この値がツールヒントとして表示されます。 この値は `description` 、定義されていない場合、またはイメージの `logoSourceAddress` ソース アドレスが破損している場合は、プレーン テキスト フォールバックとしても表示されます。

次の JSON コード スニペットは、プロパティを使用してエンティティの `provider` データ プロバイダー属性を指定するエンティティ値を示しています。

> [!NOTE]
> 完全なコード サンプルでこの JSON を使用する方法については、 [OfficeDev/office-js-snippets](https://github.com/OfficeDev/office-js-snippets/blob/main/samples/excel/85-preview-apis/data-types-entity-attribution.yaml) リポジトリを参照してください。

```TypeScript
const entity: Excel.EntityCellValue = {
    type: Excel.CellValueType.entity,
    text: productName,
    properties: {
        // Enter property settings here.
    },
    layouts: {
        // Enter layout settings here.
    },
    provider: {
        description: product.providerName, // Name of the data provider. Displays as a tooltip when hovering over the logo. Also displays as a fallback if the source address for the image is broken.
        logoSourceAddress: product.sourceAddress, // Source URL of the logo to display.
        logoTargetAddress: product.targetAddress // Destination URL that the logo navigates to when selected.
    }
};
```

次のスクリーンショットは、上記のコード スニペットを使用するエンティティ値カードを示しています。 スクリーンショットは、左下隅のデータ プロバイダーの属性を示しています。 この場合、データ プロバイダーは Microsoft であり、Microsoft ロゴが表示されます。

:::image type="content" source="../images/excel-data-types-entity-card-attribution.png" alt-text="カード レイアウト ウィンドウが表示されたエンティティ値データ型を示すスクリーンショット。カードには、左下隅にデータ プロバイダーの属性が表示されます。":::

## <a name="see-also"></a>関連項目

- [Excel アドインのデータ型の概要](excel-data-types-overview.md)
- [Excel データ型の主要概念](excel-data-types-concepts.md)
- [Excel JavaScript API リファレンス](../reference/overview/excel-add-ins-reference-overview.md)