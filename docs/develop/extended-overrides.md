---
title: マニフェストの拡張オーバーライドを処理する
description: マニフェストの拡張オーバーライドを使用して機能拡張機能を構成する方法について学習します。
ms.date: 02/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 43a922f559100157dbdacbb401d38c4d9ba22010
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743796"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a>マニフェストの拡張オーバーライドを使用する

Office アドインの一部の機能拡張機能は、アドインの XML マニフェストではなく、サーバーでホストされている JSON ファイルで構成されます。

> [!NOTE]
> この記事では、アドイン マニフェストOfficeアドインでの役割について理解している必要があります。最近Office[アドインの XML](add-in-manifests.md) マニフェストを参照してください。

次の表は、機能のドキュメントへのリンクと共に、拡張オーバーライドを必要とする機能拡張機能を指定します。

| 機能 | 開発手順 |
| :----- | :----- |
| キーボード ショートカット | [カスタム キーボード ショートカットをアドインOffice追加する](../design/keyboard-shortcuts.md) |

JSON 形式を定義するスキーマは [、拡張マニフェスト スキーマです](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json)。

> [!TIP]
> この記事はやや抽象的です。 表の記事の 1 つを読んで、概念をわかりやすくする方法を検討してください。

## <a name="tell-office-where-to-find-the-json-file"></a>JSON ファイルOffice場所を確認する

マニフェストを使用して、JSON Office場所を確認します。 マニフェスト *内* の要素の直下 ( `<VersionOverrides>` 内部ではない) に [ExtendedOverrides 要素を追加](../reference/manifest/extendedoverrides.md) します。 属性を `Url` JSON ファイルの完全な URL に設定します。 最も単純な要素の例を次に示 `<ExtendedOverrides>` します。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

次に、非常に単純な拡張オーバーライド JSON ファイルの例を示します。 これは、アドインの作業ウィンドウを開く関数 (他の場所で定義) にキーボード ショートカット Ctrl + Shift +A を割り当てる。

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

## <a name="localize-the-extended-overrides-file"></a>拡張上書きファイルをローカライズする

アドインが複数のロケール`ResourceUrl``<ExtendedOverrides>`をサポートしている場合は、要素の属性を使用して、ローカライズされたリソースOfficeをポイントできます。 次に例を示します。

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

リソース ファイルを作成して使用する方法、拡張オーバーライド ファイル内のリソースを参照する方法、およびここで説明していない追加のオプションの詳細については、「 [Localize extended overrides](localization.md#localize-extended-overrides)」を参照してください。
