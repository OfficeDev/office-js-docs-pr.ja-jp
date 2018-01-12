# <a name="office-mix-add-ins"></a>Office Mix アドイン




Microsoft Office Mix は、PowerPoint プレゼンテーションでオーディオ ビジュアル コンポーネントを有効にするために PowerPoint に組み込む Office アドインを作成するプラットフォームです。一方、LabsJS は、Office Mix に挿入できるラボと呼ばれる特殊な Office Mix インスタンスを作成するためのテクノロジです。ラボは、ユーザーがシミュレーション、デモンストレーション、クイズのような完全な対話型の教育用コンポーネントを作成できるように、Office Mix の機能を拡張します。

## <a name="lets-start-with-office-mix"></a>Office Mix の作業を開始しよう

最初に、エンド ポイントである Office Mix について説明します。Office Mix は、ビデオ、音声、音楽、インク、およびスライド上のアクションなどの教育コンポーネントを Microsoft PowerPoint プレゼンテーションに組み込んで、オンラインで発行できるようにする Microsoft のテクノロジです。ユーザーは、「ミックス」と呼ばれるこれらの動的な教育プレゼンテーションを使用して、PowerPoint プレゼンテーションを動的なレッスンにすることができます。

実行中のミックスの例については、[Office Mix ギャラリー](https://mix.office.com/Gallery)で、多数の魅力的な Office Mix のレッスンのデモを参照してください。これらの各レッスンでのミックスの効果的な使用法に注目してください。


## <a name="how-does-labsjs-fit-in-with-office-mix"></a>Office Mix との関係での LabsJS の役割

LabsJS は、Office Mix の概念を拡張するものです。Office Mix が PowerPoint プレゼンテーションを動的なレッスンにする一方で、LabsJS を使用することにより、ユーザーは拡張された方法でレッスンを操作することができます。対話機能は、デモンストレーション、クイズ、シミュレーション、およびその他多数の種類の対話型コンテンツの形で提供されます。これらの新しい、対話型の教育コンポーネントは、「ラボ」と呼ばれます。これらのラボの実体は HTML5 と JavaScript を使用して作成された単なる Office Mix アドインです。

これらのラボは、実際は labs.js API ([LabsJS JavaScript API リファレンス](../../../reference/office-mix/labsjs-javascript-api-reference.md)) を使用して JavaScript で作成されたアドインです。Labs.js は、office.js ライブラリの最上層である抽象層の役割をします。nutshell では、labs.js を使用することで、ラボを作成して Office Mix インスタンス、すなわち「ミックス」に挿入し、それらを PowerPoint で表示することが可能になります。


## <a name="take-a-look"></a>確認する

[Office Mix ギャラリー](https://mix.office.com/Gallery)は既に紹介しましたが、LabsJS ラボを含み、実行する 3 つの Office Mix の例に特に注目してください。これらの例を確認して、ラボの可能性についてのヒントを得てください。以下で、どこまでが PowerPoint で、どこからが Office Mix テクノロジか、またどこから LabsJS 機能が始まってミックスの操作が可能となるのかということに注意してください。


- [Online Python Tutor](https://mix.office.com/watch/1tkuqw9i7m4jr)
    
- [PhET 対話型シミュレーション](https://mix.office.com/watch/obibkt80fj52)
    
- [Code Hunt](https://mix.office.com/watch/q4tnp5au9mbo)
    

