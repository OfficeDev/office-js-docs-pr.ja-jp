---
ms.date: 05/03/2019
description: Excel カスタム関数をローカライズします。
title: カスタム関数をローカライズする
localization_priority: Normal
ms.openlocfilehash: 5dbe2f78f1d24c3d8c8214f4e604e66f097adba3
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628033"
---
# <a name="localize-custom-functions"></a>カスタム関数をローカライズする

アドインとカスタム関数名の両方をローカライズできます。 関数の JSON ファイルでローカライズされた関数名を指定し、XML マニフェストファイルでロケール情報を指定する必要があります。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!IMPORTANT]
> 自動生成されたメタデータはローカライズには使用できないため、JSON ファイルを手動で更新する必要があります。

## <a name="localize-function-names"></a>関数名をローカライズする

カスタム関数をローカライズするには、言語ごとに新しい JSON メタデータファイルを作成します。 各言語 JSON ファイルで、ターゲット`name`言語`description`でプロパティを作成します。 英語の既定のファイルの名前は、**関数 json**です。 追加の JSON ファイル (たとえば、**関数**の識別を容易にする) ごとに、ファイル名のロケールを使用することをお勧めします。

は`name` 、 `description` Excel に表示され、ローカライズされています。 ただし、各`id`関数のはローカライズされていません。 この`id`プロパティは、Excel で関数が一意であることを識別する方法であり、設定後に変更することはできません。

次の JSON は、"掛け算" という`id`プロパティを持つ関数を定義する方法を示しています。 関数`name`の`description`およびプロパティは、ドイツ語にローカライズされています。 各パラメーター `name`と`description`は、ドイツ語にローカライズされています。

```JSON
{
    "id": "MULTIPLY",
    "name": "SUMME",
    "description": "Summe zwei Zahlen",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "eins",
            "description": "Erste Nummer",
            "dimensionality": "scalar"
        },
        {
            "name": "zwei",
            "description": "Zweite Nummer",
            "dimensionality": "scalar"
        },
    ],
}
```

前の JSON を次の JSON と比較して英語を比較します。

```JSON
{
    "id": "MULTIPLY",
    "name": "Multiply",
    "description": "Multiplies two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "one",
            "description": "first number",
            "dimensionality": "scalar"
        },
        {
            "name": "two",
            "description": "second number",
            "dimensionality": "scalar"
        },
    ],
}
```

## <a name="localize-your-add-in"></a>アドインをローカライズする

各言語の JSON ファイルを作成した後、各 JSON メタデータファイルの URL を指定するロケールごとに、XML マニフェストファイルを上書き値で更新する必要があります。 次のマニフェスト XML は、( `en-us`ドイツ) 用の JSON ファイルの`de-de`上書き URL を含む既定のロケールを示しています。 **関数の de**ファイルには、ローカライズされたドイツ語の関数名と id が含まれています。

```XML
<DefaultLocale>en-us</DefaultLocale>
...
<Resources>
     <bt:Urls>
        <bt:Url id="Contoso.Functions.Metadata.Url" DefaultValue="https://localhost:3000/dist/functions.json"/>
          <bt:Override Locale="de-de" Value="https://localhost:3000/dist/functions-de.json" />
        </bt:url>
        
     </bt:Urls>
</Resources>
```

アドインのローカライズプロセスの詳細については、「 [Office アドインのローカライズ](../develop/localization.md#control-localization-from-the-manifest)」を参照してください。

## <a name="next-steps"></a>次の手順
[カスタム関数の名前付け規則](custom-functions-naming.md)について、または[エラー処理のベストプラクティス](custom-functions-errors.md)を検出する方法について説明します。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [カスタム関数の JSON メタデータを自動生成します](custom-functions-json-autogeneration.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
