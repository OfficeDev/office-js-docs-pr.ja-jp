---
ms.date: 04/03/2019
description: JSDOC タグを使用して、カスタム関数の JSON メタデータを動的に作成します。
title: カスタム関数の JSON メタデータを作成する (プレビュー)
localization_priority: Priority
ms.openlocfilehash: c6d89684da2d0773ccfb1763e5e3e426e647523b
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/04/2019
ms.locfileid: "31478963"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a>カスタム関数の JSON メタデータを作成する (プレビュー)

Excel カスタム関数が JavaScript または TypeScript で書き込まれている場合、JSDoc タグを使用するとカスタム関数に関する追加の情報が得られます。 JSDoc タグはビルド時に使用して、[JSON メタデータ ファイル](custom-functions-json.md)を作成します。 JSDoc タグを使用して、JSON メタデータ ファイルを手動で編集してから作業内容を保存します。

JavaScript または TypeScript 関数のコード コメントに`@customfunction`タグを追加して、カスタム関数としてマークします。

関数パラメーターの型は、JavaScript の[@param](#param)タグ、または TypeScript の[関数の型](http://www.typescriptlang.org/docs/handbook/functions.html)から指定されることがあります。 詳細については、「[@param](#param)タグと[型](#Types)」セクションを参照してください。

## <a name="jsdoc-tags"></a>JSDoc タグ
Excel カスタム関数では、次の JSDoc タグがサポートされています。
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

関数がキャンセルされた場合は、カスタム関数がアクションを実行する必要があるということです。

最後の関数パラメーターは`CustomFunctions.CancelableInvocation`の型にしてください。 関数がキャンセルされた場合、関数は`oncanceled`プロパティに関数を割り当て、実行するアクションを示すことができます。

最後の関数のパラメーターが`CustomFunctions.CancelableInvocation`の型である場合、タグが存在しない場合でも`@cancelable`と見なされます。

関数には`@cancelable`と`@streaming`のタグの両方を含めることはできません。

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

構文: @customfunction _id_ _name_

このタグを指定すると、Excel のカスタム関数として JavaScript または TypeScript の関数を処理できます。

このタグには、カスタム関数のメタデータを作成する必要があります。

次への呼び出しもあります `CustomFunctions.associate("id", functionName);`

#### <a name="id"></a>id 

この id は、文書に保存されているカスタム関数の不変の識別子として使用されます。 変更する必要はありません。

* id が提供されていない場合、JavaScript または TypeScript の関数名は大文字に変換され、許可されていない文字は削除されます。
* id はすべてのカスタム関数に一意である必要があります。
* 使用できる文字は、A～Z、a～z、0～9、ピリオド (.) に制限されています。

#### <a name="name"></a>name

カスタム関数の表示名を提供します。 

* name が指定されていない場合、id も名前として使用します。
* 使用できる文字は、文字[Unicode アルファベット](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、ピリオド (.)、アンダースコア (\_)です。
* 最初は文字にしてください。
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

JavaScript 構文: @param{type} 名_の説明_

* `{type}` 中かっこ内の型の情報を指定する必要があります。 使用する可能性のある型に関する詳細については、「[型](##types)」を参照してください。 オプション: 指定しない場合、`any`の型を使用します。
* `name` @paramのタグを適用するパラメーターを指定します。 必須です。
* `description` 関数のパラメーターの Excel で表示される説明を取得できます。 省略可能です。

オプションとしてカスタム関数のパラメーターを示すには:
* パラメーター名を角かっこで囲みます。 例: `@param {string} [text] Optional text`。

#### <a name="typescript"></a>TypeScript

TypeScript 構文: @param名前_の説明_

* `name` @paramのタグを適用するパラメーターを指定します。 必須です。
* `description` 関数のパラメーターの Excel で表示される説明を取得できます。 省略可能です。

使用する可能性のある関数のパラメーターの型に関する詳細については、「[型](##types)」を参照してください。

オプションとしてカスタム関数のパラメーターを示すには、以下のいずれかを実行します。
* オプションのパラメーターを使用します。 次に例を示します。 `function f(text?: string)`
* パラメーターに既定値を指定します。 次に例を示します。 `function f(text: string = "abc")`

@paramの詳しい説明については、「[JSDoc](http://usejsdoc.org/tags-param.html)」を参照してください。

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

機能が評価されているセルのアドレスを指定する必要があることを示しています。 

最後の関数のパラメーターは、`CustomFunctions.Invocation`の型または派生型にしてください。 関数が呼び出される場合、`address`プロパティにはアドレスが含まれます。

---
### <a name="returns"></a>@returns
<a id="returns"/>

構文: @returns {_type_}

戻り値の型を指定します。

`{type}`を省略すると、TypeScript 型の情報が使用されます。 型の情報がない場合、`any`になります。

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

カスタム関数がストリーミング関数であることを示すのに使用されます。 

最後のパラメーターは`CustomFunctions.StreamingInvocation<ResultType>`型のいずれかにする必要があります。
関数は`void`を返す必要があります。

ストリーミング関数は直接値を返しませんが、代わりに最後のパラメーターを使用して`setResult(result: ResultType)`を呼び出す必要があります。

ストリーム関数によってスローされる例外は無視されます。 `setResult()` エラー結果を示すためにエラーで呼び出されることがあります。

ストリーミング関数は[@volatile](#volatile)としてマークされます。

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

揮発性関数とは、引数を取らない場合や引数が変更されていない場合でも、結果が頻繁に同じになるとは想定できないものです。 Excel は再計算が実行される度に、揮発性関数が含まれているセルをすべての参照先と共に再評価します。 このため、揮発性関数に依存しすぎていると、再計算にかかる時間が長くなる可能性があるため、使いすぎないようにしてください。

ストリーミング関数を揮発性関数にはできません。

---

## <a name="types"></a>型

パラメーターの型を指定すると、Excel は値を指定した型に変換してから関数を呼び出します。 型が`any`の場合、変換は実行されません。

### <a name="value-types"></a>値の型

1 つの値は、`boolean`、 `number`、`string`の型のいずれかを使用して表現できる場合があります。

### <a name="matrix-type"></a>マトリックスの型

2 次元配列型を使用して、パラメーターまたは戻り値をマトリックス値にします。 たとえば、`number[][]`の型は数字のマトリックスを示します。 `string[][]` 文字列のマトリックスを示します。 

### <a name="error-type"></a>エラーの種類

非ストリーミング関数は、エラーの種類を返してエラーを示すことができます。

ストリーミング関数は、エラーの種類で setResult() を返してエラーを示すことができます。

### <a name="promise"></a>Promise

関数は Promise を返すことができ、Promise が解決される場合に値を指定します。 Promise が拒否される場合は、エラーになっています。

### <a name="other-types"></a>その他の型

その他の型はエラーとして処理されます。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [カスタム関数の変更ログ](custom-functions-changelog.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
* [カスタム関数のデバッグ](custom-functions-debugging.md)
