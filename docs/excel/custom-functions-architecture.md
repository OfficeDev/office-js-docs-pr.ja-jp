---
ms.date: 03/20/2019
description: Excelのカスタム関数のランタイムについて解説します。
title: カスタム関数のアーキテクチャ (プレビュー)
localization_priority: Priority
ms.openlocfilehash: b3f3d6c5eda51639a734c6d0f162c596f0c1e41b
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448603"
---
# <a name="custom-functions-architecture"></a>カスタム関数のアーキテクチャ

 カスタム関数は、計算の実行の優先付けを行う独自のランタイムを持っています。 この記事では、カスタム関数ランタイムと、アドインの他の部分を駆動するブラウザベースのJavaScriptエンジンの違いについて説明します。

## <a name="custom-functions-runtime"></a>カスタム関数のランタイム

Office Webアドインは、作業ウィンドウまたはコンテンツウィンドウとしてユーザーと対話したり、コマンドやカスタム機能を含めることができます。 カスタム関数を除いて、これらすべての部分はブラウザエンジンランタイムで動作します。 カスタム関数は、計算速度を最適化する別のカスタム関数の実行時に実行します。

プロジェクトの生成に [Officeアドイン用のYeomanジェネレータ](https://www.npmjs.com/package/generator-office) を使用している場合は、カスタム関数ランタイムはfunctions.htmlファイルで参照されているcustom-functions.jsスクリプトファイルを介してロードされます。 Functions.html は、ランタイムを読み込む場合にのみ機能し、アドイン用の作業ウィンドウとして使用することはできません。

次の表は、カスタム関数の実行時とブラウザーのエンジンの実行時の違いを示しています。

| カスタム関数のランタイム  | ブラウザエンジン ランタイム    |
|------------------------------------------------------------------ |-------------------------------------------------------------------------------------------------------------- |
| セルの値を返すことをサポートしています    | Office.js Api と UI 要素をサポートしています。   |
| `localStorage` オブジェクトを持たず、代わりにこちらを使用します`AsyncStorage`  | `localStorage` オブジェクトを持ち, オプションでこちらを使用することもできます`AsyncStorage`   |
| DOM の関連操作や、jQuery など DOM に依存するライブラリの読み込みはサポートされていません。    | DOM の関連操作や、DOM に依存するライブラリの読み込みがサポートされています。 |


## <a name="browser-engine-runtime"></a>ブラウザエンジン ランタイム

作業ウィンドウ、コンテンツアドイン、およびコマンドは、ブラウザエンジンランタイムで実行されます。

ブラウザエンジン ランタイムは、Office.js Api をサポートしています。 Excelのテーブルを操作できるAPIなどのExcel APIは、ブラウザエンジンランタイムで実行されますが、カスタム関数ランタイムから直接アクセスすることはできません。

## <a name="communicate-between-runtimes"></a>ランタイム間のコミュニケーション

カスタム関数のコードは、実行時間が異なるため、作業ウィンドウのようにWebアドインの他の部分のコードと直接対話することはできません。 ただし、一部のシナリオでは、トークンを渡すなどのデータを共有する必要があります。

`AsyncStorage` を、カスタム関数からのデータを保存したり、作業ウィンドウのコードからデータを取得したりするために使用できます。 データの保管と共有について詳しくは、[状態の保存と共有](custom-functions-overview.md#saving-and-sharing-state)を参照してください。

`AsyncStorage`を使用した以下のパターンとプラクティス専用の [Githubリポジトリ](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) でコードサンプルを見ることができます。
`AsyncStorage`に関する一般的な情報については、[カスタム関数ランタイム](./custom-functions-runtime.md)を参照してください。

`AsyncStorage`は認証にも役立つ場合があります。 詳細については、[カスタム関数の認証](custom-functions-authentication.md)を参照してください。

## <a name="see-also"></a>関連項目

* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [カスタム関数の変更ログ](custom-functions-changelog.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
