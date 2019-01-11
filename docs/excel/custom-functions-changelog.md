---
ms.date: 01/08/2019
description: Excel のカスタム関数に対する最新の更新内容を確認します。
title: カスタム関数の変更ログ (プレビュー)
ms.openlocfilehash: a450bf2f6a297fe0843f3b3b57a17be0c854905e
ms.sourcegitcommit: d1b5966b157a262536699eeb26ac5029cb5bf7f5
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/09/2019
ms.locfileid: "27783297"
---
# <a name="custom-functions-changelog-preview"></a>カスタム関数の変更ログ (プレビュー)

Excel カスタム関数は現時点で引き続きプレビュー段階です。つまり、変更点や新しい関数のリリースなど本製品に対して変更が頻繁に生じています。 この変更ログでは、本製品に対して加えられた変更に関する最新情報を取り上げます。

- **2017 年 11 月 7 日**: カスタム関数のプレビューとサンプルを公開*
- **2017 年 11 月 20 日**: ビルド 8801 以降を使用する場合の互換性バグを修正
- **2017 年 11 月 28 日**: 非同期関数のキャンセルのサポートを公開* (ストリーミング機能の変更が必要)
- **2018 年 5 月 7 日**: Mac、Excel Online、およびインプロセスで実行される同期関数へのサポートを公開*
- **2018 年 9 月 20日**: JavaScript ランタイムのカスタム関数へのサポートを公開。 詳細については、「[Excel カスタム関数のランタイム](custom-functions-runtime.md)」をご覧ください。
- **2018 年 10 月 20 日**: [10 月の Insider ビルド](https://support.office.com/ja-JP/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24)では、カスタム関数は、 Windows デスクトップ用およびオンライン用の[カスタム定義メタデータ](custom-functions-json.md)で 'id' パラメーターが必要になりました。 Mac では、このパラメーターは無視します。 カスタム関数では、オプションのパラメーターおよび `any` 戻り値の型もサポートされるようになりました。
- **2018 年 12 月 12 日**: カスタム関数にセル アドレスを検索する手段が備わりました。 詳しくは、「[カスタム関数が呼び出したセルを特定する](custom-functions-overview.md#determine-which-cell-invoked-your-custom-function)」をご覧ください。
- **2019 年 1 月 8 日**: バインド メソッド `CustomFunctionMapping()` が `CustomFunctions.associate()` に変更されました。 詳細については、「[カスタム関数のベスト プラクティス](custom-functions-best-practices.md)」を参照してください。

\* [Office Insider](https://products.office.com/office-insider) チャンネル (旧称 "Insider Fast") に対して

製品の既知の問題の一覧については、「[既知の問題](custom-functions-overview.md#known-issues)」をご覧ください。 

## <a name="see-also"></a>関連項目

* [カスタム関数の概要](custom-functions-overview.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [チュートリアル: Excel でカスタム関数を作成します。](../tutorials/excel-tutorial-create-custom-functions.md)
