---
ms.date: 07/10/2019
description: Excelのカスタム関数のランタイムについて解説します。
title: カスタム関数のアーキテクチャ
localization_priority: Normal
ms.openlocfilehash: ced62f7efb826862eee8079a66fa657ea466e4b3
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950355"
---
# <a name="custom-functions-architecture"></a>カスタム関数のアーキテクチャ

 カスタム関数は、計算の実行の優先付けを行う独自のランタイムを持っています。 この記事では、カスタム関数ランタイムと、アドインの他の部分を駆動するブラウザベースのJavaScriptエンジンの違いについて説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="custom-functions-runtime"></a>カスタム関数のランタイム

Office Webアドインは、作業ウィンドウまたはコンテンツウィンドウとしてユーザーと対話したり、コマンドやカスタム機能を含めることができます。 カスタム関数を除いて、これらすべての部分はブラウザエンジンランタイムで動作します。 カスタム関数は、計算速度を最適化する別のカスタム関数の実行時に実行します。

プロジェクトの生成に [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用している場合は、カスタム関数ランタイムは **functions.html** ファイルで参照されている custom-functions.js スクリプト ファイルを介して読み込みます。 **functions.html** は、ランタイムを読み込む場合にのみ機能し、アドイン用の作業ウィンドウとして使用することはできません。

次の表は、カスタム関数の実行時とブラウザーのエンジンの実行時の違いを示しています。

| カスタム関数のランタイム  | ブラウザエンジン ランタイム    |
|------------------------------------------------------------------ |-------------------------------------------------------------------------------------------------------------- |
| セルの値を返すことをサポートしています    | Office.js Api と UI 要素をサポートしています。   |
| `localStorage` オブジェクトを持たず、代わりに `OfficeRuntime.storage` オブジェクトを使用します。     | `localStorage` オブジェクトを持ち, オプションで `OfficeRuntime.storage` オブジェクトを使用することもできます。     |
| DOM の関連操作や、jQuery など DOM に依存するライブラリの読み込みはサポートされていません。    | DOM の関連操作や、DOM に依存するライブラリの読み込みがサポートされています。 |

## <a name="browser-engine-runtime"></a>ブラウザエンジン ランタイム

作業ウィンドウ、コンテンツアドイン、およびコマンドは、ブラウザエンジンランタイムで実行されます。

ブラウザエンジン ランタイムは、Office.js Api をサポートしています。 Excelのテーブルを操作できるAPIなどのExcel APIは、ブラウザエンジンランタイムで実行されますが、カスタム関数ランタイムから直接アクセスすることはできません。

## <a name="communicate-between-runtimes"></a>ランタイム間のコミュニケーション

カスタム関数のコードは、実行時間が異なるため、作業ウィンドウのようにWebアドインの他の部分のコードと直接対話することはできません。 ただし、一部のシナリオでは、トークンを渡すなどのデータを共有する必要があります。

`OfficeRuntime.storage` オブジェクトを、カスタム関数からのデータを保存したり、作業ウィンドウのコードからデータを取得したりするために使用できます。 データの保管と共有の詳細については、「[状態の保存と共有](custom-functions-save-state.md)」を参照してください。

パターンとプラクティス専用の [Githubリポジトリ](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) で `storage` オブジェクトを使用してコード サンプルを見ることができます。
`storage` オブジェクトに関する一般的な情報の詳細については、「[カスタム関数ランタイム](./custom-functions-runtime.md)」を参照してください。

`storage` オブジェクトは認証にも役立つ場合があります。 詳細については、[カスタム関数の認証](custom-functions-authentication.md)を参照してください。

## <a name="next-steps"></a>次の手順
詳細については、「[カスタム関数ランタイムの使用](custom-functions-runtime.md)」を参照してください。

## <a name="see-also"></a>関連項目

* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数でデータを受信して​​処理する](custom-functions-web-reqs.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
