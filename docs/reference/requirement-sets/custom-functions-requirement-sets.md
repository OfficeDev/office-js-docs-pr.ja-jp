---
title: カスタム関数の要件セット
description: JavaScript API のカスタム関数要件セットExcel詳細です。
ms.date: 02/15/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 3f8e66932d6f960898dc2185299fcbe58324f281
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746220"
---
# <a name="custom-functions-requirement-sets"></a>カスタム関数の要件セット

[カスタム関数](../../excel/custom-functions-overview.md)は、コア Excel JavaScript API の個別の要件セットを使用します。 次の表に、カスタム関数の要件セット、サポートされているクライアント アプリケーションOffice、それらのアプリケーションのビルド バージョンまたは番号を示します。

|  要件セット  |  Office 2021 以降のWindows<br>(1 回限りの購入)  |  Windows での Office<br>(Microsoft 365 サブスクリプションに接続)  |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(両方のサブスクリプション<br> Mac 2021 以降Office 1 回購入)  | Office on the web |
|:-----|:-----|:-----|:-----|:-----|:-----|
| CustomFunctionsRuntime 1.3 | 16.0.14326.20454 以降 | 16.0.13127.20296 以降 | サポート対象外 | 16.40.20081000 以降 | 2020 年 7 月 |
| CustomFunctionsRuntime 1.2 | 16.0.14326.20454 以降 | 16.0.12527.20194 以降 | サポート対象外 | 16.34.20020900 以降 | 2020 年 1 月 |
| CustomFunctionsRuntime 1.1 | 16.0.14326.20454 以降 | 16.0.12527.20092 以降 | サポート対象外 | 16.34 以降 | 2019 年 5 月 |

## <a name="customfunctionsruntime-11-12-and-13"></a>CustomFunctionsRuntime 1.1、1.2、および 1.3

CustomFunctionsRuntime 1.1 は API の最初のバージョンです。 要件セット 1.2 は、エラー処理をサポート `CustomFunctions.Error` するオブジェクトを追加します。 要件セット 1.3 は [、XLL ストリーミング](../../excel/make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) `ErrorCode` サポートと新しいオプションを [CustomFunctions.Error オブジェクトに追加](/javascript/api/custom-functions-runtime/customfunctions.error) します。

## <a name="see-also"></a>関連項目

- [カスタム関数リファレンス ドキュメント](/javascript/api/custom-functions-runtime)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
