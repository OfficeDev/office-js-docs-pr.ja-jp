---
title: カスタム関数の要件セット
description: JavaScript API のカスタム関数要件セットExcel詳細です。
ms.date: 09/08/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: af35f88041f5c90268782fb4cee44b44b56c4644
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/18/2021
ms.locfileid: "59445852"
---
# <a name="custom-functions-requirement-sets"></a>カスタム関数の要件セット

[カスタム関数](../../excel/custom-functions-overview.md)は、コア Excel JavaScript API の個別の要件セットを使用します。 次の表に、カスタム関数の要件セット、サポートされているクライアント アプリケーションOffice、それらのアプリケーションのビルド バージョンまたは番号を示します。

|  要件セット  |  Office 2021 以降のWindows<br>(1 回限りの購入)  |  Windows での Office<br>(Microsoft 365 サブスクリプションに接続)  |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web |
|:-----|:-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | 16.0.13127.20296 以降 | 16.0.13127.20296 以降 | サポート対象外 | 16.40.20081000 以降 | 2020 年7 月 |
| CustomFunctionsRuntime 1.2 | 16.0.13127.20296 以降 | 16.0.12527.20194 以降 | 非サポート | 16.34.20020900 以降 | 2020 年 1 月 |
| CustomFunctionsRuntime 1.1 | 16.0.13127.20296 以降 | 16.0.12527.20092 以降 | 非サポート | 16.34 以降 | 2019 年 5 月 |

> [!NOTE]
> Excelカスタム関数は、2019 Office以前 (1 回の購入) ではサポートされていません。

## <a name="customfunctionsruntime-11-12-and-13"></a>CustomFunctionsRuntime 1.1、1.2、および 1.3

CustomFunctionsRuntime 1.1 は API の最初のバージョンです。 要件セット 1.2 は、エラー処理 `CustomFunctions.Error` をサポートするオブジェクトを追加します。 要件セット 1.3 は [、XLL ストリーミング](../../excel/make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) サポートと新しいオプションを `ErrorCode` [CustomFunctions.Error オブジェクトに追加](/javascript/api/custom-functions-runtime/customfunctions.error) します。

## <a name="see-also"></a>関連項目

- [カスタム関数リファレンス ドキュメント](/javascript/api/custom-functions-runtime)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
