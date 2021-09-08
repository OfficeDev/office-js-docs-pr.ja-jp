---
ms.date: 07/08/2021
description: JSDoc タグを使用して、カスタム関数の JSON メタデータを動的に作成します。
title: カスタム関数用の JSON メタデータの自動生成
localization_priority: Normal
ms.openlocfilehash: b4ae61ab46de7dadb9280e731d65715adaf64630
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936565"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a>カスタム関数用の JSON メタデータの自動生成

Excel カスタム関数が JavaScript または TypeScript で記述されている場合、カスタム関数に関する追加の情報を提供するために、[JSDoc タグ](https://jsdoc.app/)が使用されます。 JSDoc タグはビルド時に使用して、JSON メタデータ ファイルを作成します。 JSDoc タグを使用すると、JSON メタデータ ファイルを手動で [編集する手間が省きます](custom-functions-json.md)。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

JavaScript または TypeScript 関数のコード コメントに`@customfunction`タグを追加して、カスタム関数としてマークします。

関数パラメーターの型は、JavaScript の [@param](#param) タグを使用して指定するか、TypeScript の[関数の型](https://www.typescriptlang.org/docs/handbook/functions.html)から指定できます。 詳細については、「[@param](#param) タグ」セクションと「[型](#types)」セクションを参照してください。

### <a name="add-a-description-to-a-function"></a>関数に説明を追加する

説明は、カスタム関数の機能を理解するためのヘルプが必要な場合に、ヘルプ テキストとしてユーザーに表示されます。 説明に特定のタグは必要ありません。 JSDoc コメントに簡単な説明を入力するだけです。 一般に、説明は JSDoc コメント セクションの先頭に配置されますが、配置場所に関係なく機能します。

組み込み関数の説明の例を表示するには、Excel を開き、**[数式]** タブに移動し、**[関数の​​挿入]** を選択します。 すべての関数の説明を参照したり、独自のカスタム関数を一覧表示したりすることができます。

次の例では、「球の体積を計算します。」 が、カスタム関数の説明です。

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```

## <a name="jsdoc-tags"></a>JSDoc タグ

次の JSDoc タグは、カスタム関数Excelサポートされています。

* [@cancelable](#cancelable)
* [@customfunction](#customfunction) id 名
* [@helpurl](#helpurl) url
* [@param](#param) _{type}_ 名前の説明
* [@requiresAddress](#requiresAddress)
* [@requiresParameterAddresses](#requiresParameterAddresses)
* [@returns](#returns) _{type}_
* [@streaming](#streaming)
* [@volatile](#volatile)

---
<a id="cancelable"></a>
### <a name="cancelable"></a>@cancelable

関数が取り消された場合に、カスタム関数がアクションを実行します。

最後の関数パラメーターは `CustomFunctions.CancelableInvocation` の型にする必要があります。 関数はプロパティに関数を割り当て、関数が取り消された場合の結果 `oncanceled` を示します。

最後の関数のパラメーターが `CustomFunctions.CancelableInvocation` 型の場合、タグは表示されませんが、`@cancelable` と見なされます。

関数には `@cancelable` と `@streaming` の両方のタグを含めることはできません。

<a id="customfunction"></a>

### <a name="customfunction"></a>@customfunction

構文: @customfunction _id_ _名_

このタグは、JavaScript/TypeScript 関数がカスタム関数Excel示します。 カスタム関数のメタデータを作成する必要があります。

次に、このタグの例を示します。

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a>id

カスタム `id` 関数を識別します。

* `id` が提供されていない場合、JavaScript または TypeScript の関数名は大文字に変換され、許可されない文字は削除されます。
* `id` はすべてのカスタム関数で一意である必要があります。
* 指定できる文字は、A から Z、a から z、0 から 9、アンダースコア (\_)、ピリオド (.) に制限されます。

次の例では、インクリメントは関数の `id` と `name` です。

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a>name

カスタム関数の表示用の `name` を提供します。

* name が指定されていない場合、id が名前としても使用されます。
* 使用できる文字は、文字 [Unicode アルファベット](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic)、数字、ピリオド (.)、およびアンダースコア (\_)です。
* 最初の文字は、アルファベット文字にする必要があります。
* 最大文字数は 128 文字です。

次の例では、INC は関数の`id` で、 `increment` は`name`です。

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a>説明

説明は、関数の入力中Excelユーザーに表示され、関数の動作を指定します。 説明に特定のタグは必要ありません。 JSDoc コメント内に関数の機能を説明するフレーズを入力して、カスタム関数に説明を追加します。 既定では、JSDoc コメント セクションでタグが付けられていないテキストは、関数の説明です。

次の例では、「2 つの数値を加算する関数」というフレーズが、ID プロパティ `ADD` のカスタム関数の説明です。

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

<a id="helpurl"></a>
### <a name="helpurl"></a>@helpurl

構文: @helpurl _url_

指定された _url_ が Excel で表示されます。

次の例では、 `helpurl` です `www.contoso.com/weatherhelp` 。

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

<a id="param"></a>
### <a name="param"></a>@param

#### <a name="javascript"></a>JavaScript

JavaScript 構文: @param {type} 名 _の説明_

* `{type}` 中かっこ内の型情報を指定します。 使用できる型に関する詳細については、「[型](#types)」セクションを参照してください。 型が指定されていない場合は、既定の型 `any` が使用されます。
* `name` タグが適用されるパラメーター@param指定します。 必須です。
* `description` は、Excel で表示される関数のパラメーターの説明を示します。 オプションです。

カスタム関数パラメーターを省略可能として指定するには、パラメーター名の周囲に角かっこを付けます。 たとえば、`@param {string} [text] Optional text` などです。

> [!NOTE]
> 省略可能なパラメーターの既定値は `null` です。

次の例は、2 つまたは 3 つの数値を追加する ADD 関数を示しています。3 番目の数値は省略可能なパラメーターです。

```js
/**
 * A function which sums two, or optionally three, numbers.
 * @customfunction ADDNUMBERS
 * @param firstNumber {number} First number to add.
 * @param secondNumber {number} Second number to add.
 * @param [thirdNumber] {number} Optional third number you wish to add.
 * ...
 */
```

#### <a name="typescript"></a>TypeScript

TypeScript 構文: @param 名 _の説明_

* `name` タグが適用されるパラメーター@param指定します。 必須です。
* `description` は、Excel で表示される関数のパラメーターの説明を示します。 オプションです。

使用できる関数のパラメーターの型に関する詳細については、「[型](#types)」セクションを参照してください。

カスタム関数のパラメーターを省略可能として示すには、以下のいずれかを実行します。

* 省略可能なパラメーターを使用する。 例: `function f(text?: string)`
* パラメーターに既定値を指定する。 例: `function f(text: string = "abc")`

@param の詳しい説明については、「[JSDoc](https://jsdoc.app/tags-param.html)」を参照してください。

> [!NOTE]
> 省略可能なパラメーターの既定値は `null` です。

次の例は、2 つの数値を加算する `add` 関数を示しています。

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

<a id="requiresAddress"></a>

### <a name="requiresaddress"></a>@requiresAddress

関数が評価されているセルのアドレスを指定する必要があることを示します。

最後の関数パラメーターは、使用する型 `CustomFunctions.Invocation` または派生型である必要があります `@requiresAddress` 。 関数が呼び出されると、`address` プロパティにアドレスが含まれます。

次のサンプルは、パラメーターを組み合わせて使用して、カスタム関数を呼び出したセルのアドレス `invocation` `@requiresAddress` を返す方法を示しています。 詳細については [、「呼び出しパラメーター](custom-functions-parameter-options.md#invocation-parameter) 」を参照してください。

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

<a id="requiresParameterAddresses"></a>
### <a name="requiresparameteraddresses"></a>@requiresParameterAddresses

関数が入力パラメーターのアドレスを返す必要があります。 

最後の関数パラメーターは、使用する型 `CustomFunctions.Invocation` または派生型である必要があります  `@requiresParameterAddresses` 。 JSDoc コメントには、戻り値を行列として指定するタグも含 `@returns` める `@returns {string[][]}` 必要があります `@returns {number[][]}` 。 詳細については [、「Matrix 型](#matrix-type) 」を参照してください。 

関数が呼び出された場合、 `parameterAddresses` プロパティには入力パラメーターのアドレスが含まれます。

次のサンプルは、3 つの入力パラメーターのアドレスを返す場合と組み合わせてパラメーターを使用 `invocation` `@requiresParameterAddresses` する方法を示しています。 詳細 [については、「パラメーターのアドレスを検出する](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) 」を参照してください。 

```js
/**
 * Return the addresses of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns {string[][]} The addresses of the parameters, as a 2-dimensional array.
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

<a id="returns"></a>
### <a name="returns"></a>@returns

構文: @returns {_type_}

戻り値の型を指定します。

`{type}` を省略すると、TypeScript の型情報が使用されます。 型情報がない場合、型は `any` になります。

次の例は、 `@returns` タグを使用する `add` 関数を示しています。

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

<a id="streaming"></a>
### <a name="streaming"></a>@streaming

カスタム関数がストリーミング関数であることを示すのに使用されます。 

最後のパラメーターは型です `CustomFunctions.StreamingInvocation<ResultType>` 。
この関数は、 を返します `void` 。

ストリーミング関数は値を直接返すのではなく、最後のパラメーターを使用 `setResult(result: ResultType)` して呼び出します。

ストリーム関数によってスローされる例外は無視されます。 `setResult()` が、エラー結果を示すために、Error により呼び出されることがあります。 ストリーミング関数と詳細については、「[ストリーミング関数を作成する](custom-functions-web-reqs.md#make-a-streaming-function)」を参照してください。

ストリーミング関数は、[@volatile](#volatile) としてマークできません。

<a id="volatile"></a>
### <a name="volatile"></a>@volatile

揮発性関数とは、引数を取らない場合や引数が変更されていない場合でも、ある瞬間と次の瞬間では結果が異なる関数です。 Excel では、再計算が実行される度に、揮発性関数を含むセルはすべての参照先と共に、再評価されます。 このため、揮発性関数を多用し過ぎると再計算にかかる時間が長くなる可能性があるため、多用しないようにします。

ストリーミング関数に揮発性関数は使用できません。

次の関数は揮発性で、 `@volatile` タグを使用します。

```js
/**
 * Simulates rolling a 6-sided die.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a>型

パラメーターの型を指定すると、Excel は値を指定した型に変換してから関数を呼び出します。 型が`any`の場合、変換は実行されません。

### <a name="value-types"></a>値の型

1 つの値は、`boolean`、 `number`、`string`の型のいずれかを使用して表現できます。

### <a name="matrix-type"></a>マトリックス型

2 次元配列型を使用して、パラメーターまたは戻り値を値のマトリックスにします。 たとえば、型は `number[][]` 数値の行列を示し、 `string[][]` 文字列の行列を示します。

### <a name="error-type"></a>エラーの種類

非ストリーミング関数は、エラーの種類を返すことによりエラーを示すことができます。

ストリーミング関数は、エラーの種類で `setResult()` を返してエラーを示すことができます。

### <a name="promise"></a>Promise

カスタム関数は、約束が解決された場合に値を提供する約束を返します。 約束が拒否された場合、カスタム関数はエラーをスローします。

### <a name="other-types"></a>その他の型

その他の型は、エラーとして処理されます。

## <a name="next-steps"></a>次の手順

[カスタム関数用の命名規則](custom-functions-naming.md)について説明します。 または、[JSON ファイルを手で書く](custom-functions-json.md)必要のある[機能をローカライズする](custom-functions-localize.md)方法を確認してください。

## <a name="see-also"></a>関連項目

* [カスタム関数の JSON メタデータを手動で作成する](custom-functions-json.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
