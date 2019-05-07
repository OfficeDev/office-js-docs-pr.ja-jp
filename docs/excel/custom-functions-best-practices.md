---
ms.date: 05/06/2019
description: Excel のカスタム関数を開発する際のベスト プラクティスについて説明します。
title: カスタム関数のベスト プラクティス
localization_priority: Normal
ms.openlocfilehash: 7369faa463966dd309258bf431eae8719407be38
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628145"
---
# <a name="custom-functions-best-practices"></a>カスタム関数のベスト プラクティス

この記事では、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="associating-function-names-with-json-metadata"></a>関数名を JSON メタデータに関連付ける

[カスタム関数の概要](custom-functions-overview.md)という記事で取り上げたように、カスタム関数プロジェクトには、カスタム関数を作成するために、JSON メタデータ ファイルとスクリプト (JavaScript または TypeScript) の両方を含める必要があります。 JSON メタデータを`yo office`使用している場合は、コードコメントから生成することができます。 それ以外の場合は、JSON メタデータファイルを手動でビルドする必要があります。

関数が正しく動作するには、関数の`id`プロパティを JavaScript 実装に関連付ける必要があります。 関連付けがあることを確認してください。それ以外の場合は、関数は呼び出されません。 次のコードサンプルは、 `CustomFunctions.associate()`メソッドを使用して関連付けを行う方法を示しています。 このサンプルではカスタム関数 `add` を定義し、それを `id` プロパティ値が **ADD** の、JSON メタデータ ファイル内のオブジェクトに関連付けます。

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

次の JSON は、以前のカスタム関数 JavaScript コードに関連付けられている JSON メタデータを示しています。

```json
{
  "functions": [
    {
        "description": "Add two numbers",
        "id": "ADD",
        "name": "ADD",
        "parameters": [
            {
                "description": "First number",
                "name": "first",
                "type": "number"
            },
            {
                "description": "Second number",
                "name": "second",
                "type": "number"
            }
        ],
        "result": {
            "type": "number"
        }
    },
  ]
}
```


JavaScript ファイルでカスタム関数を作成し、JSON のメタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。

* JSON のメタデータ ファイルにそれぞれの `id` プロパティには、英数字とピリオドのみが含まれています。

* JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。 すなわち、メタデータ ファイル内の 2 つの関数オブジェクトは同じ `id` 値であってはいけません。

* 対応する JavaScript 関数の名前に関連付けられた後では、JSON のメタデータ ファイル内の `id` プロパティの値を変更しないでください。 JSON のメタデータ ファイル内の `name` プロパティを更新することによって Excel でエンド ユーザーに表示される関数の名前を変更することができます。しかし、確立された後は、 `id` プロパティの値を決して変更しないでください。

* JavaScript ファイルで、各関数の`CustomFunctions.associate`後に、カスタム関数の関連付けを指定します。

次のサンプルは、JavaScript コード サンプルで定義された関数に対応する JSON メタデータを示します。 プロパティ`id`と`name`プロパティの値は、大文字で記述します。これは、カスタム関数を記述するときのベストプラクティスです。 この JSON を追加する必要があるのは、自動生成を使用せずに、手動で独自の JSON ファイルを準備する場合だけです。 Autogeneration の詳細については、「[カスタム関数の JSON メタデータを作成](custom-functions-json-autogeneration.md)する」を参照してください。

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="additional-considerations"></a>その他の考慮事項

カスタム関数から直接または間接的に (たとえば、jQuery を使用して) ドキュメントオブジェクトモデル (DOM) にアクセスしないようにします。 カスタム関数が [JavaScript ランタイム](custom-functions-runtime.md)を使用する Excel for Windows では、カスタム関数はDOM にアクセスできません。

## <a name="next-steps"></a>次の手順
[カスタム関数を使用して web 要求を実行](custom-functions-web-reqs.md)する方法について説明します。

## <a name="see-also"></a>関連項目

* [カスタム関数の JSON メタデータを自動生成します](custom-functions-json-autogeneration.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
