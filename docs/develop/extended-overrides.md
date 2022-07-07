---
title: マニフェストの拡張オーバーライドを操作する
description: マニフェストの拡張オーバーライドを使用して機能拡張機能を構成する方法について説明します。
ms.date: 02/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 43e9820f54f2812130f7f86529c52b20b92811a0
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659953"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a>マニフェストの拡張オーバーライドを操作する

Office アドインの一部の機能拡張機能は、アドインの XML マニフェストではなく、サーバーでホストされている JSON ファイルで構成されます。

> [!NOTE]
> この記事では、Office アドイン マニフェストとアドインでの役割について理解していることを前提としています。最近ない場合 [は、Office アドインの XML マニフェスト](add-in-manifests.md)をお読みください。

次の表では、拡張機能のオーバーライドを必要とする機能拡張機能と、機能のドキュメントへのリンクを示します。

| 機能 | 開発手順 |
| :----- | :----- |
| キーボード ショートカット | [Office アドインにカスタム キーボード ショートカットを追加する](../design/keyboard-shortcuts.md) |

JSON 形式を定義するスキーマは、 [拡張マニフェスト スキーマ](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)です。

> [!TIP]
> この記事はやや抽象的です。 表の記事のいずれかを読んで、概念を明確にすることを検討してください。

## <a name="tell-office-where-to-find-the-json-file"></a>JSON ファイルを検索する場所を Office に指示する

マニフェストを使用して、JSON ファイルを見つける場所を Office に指示します。 マニフェスト内の要素の直 *下* (内部ではない) **\<VersionOverrides\>** に [ExtendedOverrides](/javascript/api/manifest/extendedoverrides) 要素を追加します。 属性を `Url` JSON ファイルの完全な URL に設定します。 可能な最も **\<ExtendedOverrides\>** 単純な要素の例を次に示します。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

非常に単純な拡張オーバーライド JSON ファイルの例を次に示します。 アドインの作業ウィンドウを開く関数 (他の場所で定義) にキーボード ショートカット Ctrl + Shift + A を割り当てます。

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+A"
            }
        }
    ]
}
```

## <a name="localize-the-extended-overrides-file"></a>拡張オーバーライド ファイルをローカライズする

アドインで複数のロケールがサポートされている場合は、要素の属性を`ResourceUrl`**\<ExtendedOverrides\>** 使用して、ローカライズされたリソースのファイルを Office にポイントできます。 次に例を示します。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

リソース ファイルを作成して使用する方法の詳細、拡張オーバーライド ファイル内のリソースを参照する方法、およびここで説明しないその他のオプションについては、「 [拡張オーバーライドのローカライズ](localization.md#localize-extended-overrides)」を参照してください。
