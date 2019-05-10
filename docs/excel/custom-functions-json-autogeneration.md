---
ms.date: 05/03/2019
description: JSDOC タグを使用して、カスタム関数の JSON メタデータを動的に作成する。
title: カスタム関数用の JSON メタデータの自動生成
localization_priority: Priority
ms.openlocfilehash: df1c0114597e2aa98a15db48c515469fb9db6cd9
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628089"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a>カスタム関数用の JSON メタデータの自動生成

Excel カスタム関数が JavaScript または TypeScript で記述されている場合、カスタム関数に関する追加の情報を提供するために、JSDoc タグが使用されます。 JSDoc タグはビルド時に使用して、[JSON メタデータ ファイル](custom-functions-json.md)を作成します。 JSDoc タグを使用すると、JSON メタデータ ファイルを手動で編集する手間が省けます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

JavaScript または TypeScript 関数のコード コメントに`@customfunction`タグを追加して、カスタム関数としてマークします。

関数パラメーターの型は、JavaScript の [@param](#param) タグを使用して指定するか、TypeScript の[関数の型](https://www.typescriptlang.org/docs/handbook/functions.html)から指定できます。 詳細については、「[@param](#param) タグ」セクションと「[型](#types)」セクションを参照してください。

## <a name="jsdoc-tags"></a>JSDoc タグ
Excel カスタム関数では、次の JSDoc タグを利用できます。
* [@cancelable](#cancelable)
* [@customfunction](#customfunction) id 名
* [@helpurl](#helpurl) url
* [@param](#param) _{type}_ 名前の説明
* [@requiresAddress](#requiresAddress)
* [@returns](#returns) _{type}_
* [@streaming](#streaming)
* [@volatile](#volatile)

---
### <a name="cancelable"></a>@cancelable
<a id="cancelable"/>

関数がキャンセルされた場合にカスタム関数がアクションを実行することを示します。

最後の関数パラメーターは `CustomFunctions.CancelableInvocation` の型にする必要があります。 関数は `oncanceled` プロパティに関数を割り当て、関数がキャンセルされた場合に実行するアクションを示すことができます。

最後の関数のパラメーターが `CustomFunctions.CancelableInvocation` 型の場合、タグは表示されませんが、`@cancelable` と見なされます。

関数には `@cancelable` と `@streaming` の両方のタグを含めることはできません。

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

構文: @customfunction _id_ _名_

このタグを指定すると、JavaScript または TypeScript の関数を、Excel のカスタム関数として処理できます。

このタグは、カスタム関数のメタデータを作成するために必要です。

次への呼び出しもあります: `CustomFunctions.associate("id", functionName);`

#### <a name="id"></a>id

id は、文書に格納されているカスタム関数の不変の識別子として使用されます。 変更する必要はありません。

* id が提供されていない場合、JavaScript または TypeScript の関数名は大文字に変換され、許可されない文字は削除されます。
* id はすべてのカスタム関数で一意である必要があります。
* 使用できる文字は、A～Z、a～z、0～9、ピリオド (.) に制限されています。

#### <a name="name"></a>name

カスタム関数の表示名を提供します。

* name が指定されていない場合、id が名前としても使用されます。
* 使用できる文字は、文字 [Unicode アルファベット](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、ピリオド (.)、およびアンダースコア (\_)です。
* 最初の文字は、アルファベット文字にする必要があります。
* 最大文字数は 128 文字です。

---
### <a name="helpurl"></a>@helpurl
<a id="helpurl"/>

構文: @helpurl _url_

指定された _url_ が Excel で表示されます。

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a>JavaScript

JavaScript 構文: @param {type} 名_の説明_

* `{type}` は、中かっこ内の型の情報を指定します。 使用できる型に関する詳細については、「[型](##types)」を参照してください。 省略可能: 指定しない場合、`any` 型が使用されます。
* `name` は、@param タグを適用するパラメーターを指定します。 必須です。
* `description` は、Excel で表示される関数のパラメーターの説明を示します。 省略可能です。

カスタム関数内のパラメーターを省略可能と指定する方法:
* パラメーター名を角かっこで囲みます。 例: `@param {string} [text] Optional text`。

> [!NOTE]
> 省略可能なパラメーターの既定値は `null` です。

#### <a name="typescript"></a>TypeScript

TypeScript 構文: @param 名 _の説明_

* `name` は、@param タグを適用するパラメーターを指定します。 必須です。
* `description` は、Excel で表示される関数のパラメーターの説明を示します。 省略可能です。

使用できる関数のパラメーターの型に関する詳細については、「[型](##types)」を参照してください。

カスタム関数のパラメーターを省略可能として示すには、以下のいずれかを実行します。
* 省略可能なパラメーターを使用する。 例: `function f(text?: string)`
* パラメーターに既定値を指定する。 例: `function f(text: string = "abc")`

@param の詳しい説明については、「[JSDoc](https://usejsdoc.org/tags-param.html)」を参照してください。

> [!NOTE]
> 省略可能なパラメーターの既定値は `null` です。

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

関数が評価されているセルのアドレスを指定する必要があることを示します。 

最後の関数のパラメーターは、`CustomFunctions.Invocation` 型または派生型にする必要があります。 関数が呼び出されると、`address` プロパティにアドレスが含まれます。

---
### <a name="returns"></a>@returns
<a id="returns"/>

構文: @returns {_type_}

戻り値の型を指定します。

`{type}` を省略すると、TypeScript の型情報が使用されます。 型情報がない場合、型は `any` になります。

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

カスタム関数がストリーミング関数であることを示すのに使用されます。 

最後のパラメーターは、`CustomFunctions.StreamingInvocation<ResultType>` 型にする必要があります。
関数は `void` を返します。

ストリーミング関数は値を直接返さず、代わりに、最後のパラメーターを使用して `setResult(result: ResultType)` を呼び出します。

ストリーム関数によってスローされる例外は無視されます。 `setResult()` が、エラー結果を示すために、Error により呼び出されることがあります。

ストリーミング関数は、[@volatile](#volatile) としてマークできません。

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

揮発性関数とは、引数を取らない場合や引数が変更されていない場合でも、ある瞬間と次の瞬間では結果が異なる可能性があると見なされる関数です。 Excel では、再計算が実行される度に、揮発性関数を含むセルはすべての参照先と共に、再評価されます。 このため、揮発性関数を多用し過ぎると再計算にかかる時間が長くなる可能性があるため、多用しないようにします。

ストリーミング関数に揮発性関数は使用できません。

---

## <a name="types"></a>型

パラメーターの型を指定すると、Excel は値を指定した型に変換してから関数を呼び出します。 型が`any`の場合、変換は実行されません。

### <a name="value-types"></a>値の型

1 つの値は、`boolean`、 `number`、`string`の型のいずれかを使用して表現できます。

### <a name="matrix-type"></a>マトリックス型

2 次元配列型を使用して、パラメーターまたは戻り値を値のマトリックスにします。 たとえば、`number[][]`の型は数字のマトリックスを示します。 `string[][]` は、文字列のマトリックスを示します。 

### <a name="error-type"></a>エラーの種類

非ストリーミング関数は、エラーの種類を返すことによりエラーを示すことができます。

ストリーミング関数は、エラーの種類で setResult() を返してエラーを示すことができます。

### <a name="promise"></a>Promise

関数は Promise を返すことができ、Promise が解決されたときに値を提供します。 Promise が拒否された場合は、エラーになります。

### <a name="other-types"></a>その他の型

その他の型は、エラーとして処理されます。

## <a name="next-steps"></a>次の手順
[カスタム関数用の命名規則](custom-functions-naming.md)について説明します。 または、[JSON ファイルを手で書く](custom-functions-json.md)必要のある[機能をローカライズする](custom-functions-localize.md)方法を確認してください。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
