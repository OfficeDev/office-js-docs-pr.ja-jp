---
title: カスタム関数の要件セット
description: Excel JavaScript API のカスタム関数要件セットの詳細。
ms.date: 09/14/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0860dd2d1b55376a85eadf04898d288d83b0205d
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819526"
---
# <a name="custom-functions-requirement-sets"></a>カスタム関数の要件セット

[カスタム関数](custom-functions-overview.md)は、コア Excel JavaScript API の個別の要件セットを使用します。 次の表に、カスタム関数要件セット、サポート対象の Office クライアントアプリケーション、およびそれらのアプリケーションのビルドバージョンまたは番号を示します。

|  要件セット  |  Windows での Office<br>(Microsoft 365 サブスクリプションに接続)  |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web |
|:-----|-----|:-----|:-----|:-----|:-----|
| Customのランタイム1.3 | 16.0.13127.20296 以降 | 非サポート | 16.40.20081000 以降 | 2020 年 7 月 |
| Customのランタイム1.2 | 16.0.12527.20194 以降 | 非サポート | 16.34.20020900 以降 | 2020 年 1 月 |
| CustomFunctionsRuntime 1.1 | 16.0.12527.20092 以降 | 非サポート | 16.34 以降 | 2019 年 5 月 |

> [!NOTE]
> Excel カスタム関数は Office 2019 またはそれ以前のバージョンではサポートされていません (1 回限りの購入)。

## <a name="customfunctionsruntime-11-12-and-13"></a>Customなランタイム1.1、1.2、1.3

Customのランタイム1.1 は、API の最初のバージョンです。 要件セット1.2 は、 `CustomFunctions.Error` エラー処理をサポートするオブジェクトを追加します。 要件セット1.3 は、 [XLL のストリーミング](make-custom-functions-compatible-with-xll-udf.md#custom-function-behavior-for-xll-compatible-functions) サポートと新しい `ErrorCode` オプションを customfunctions に追加し [ます。 Error](/javascript/api/custom-functions-runtime/customfunctions.error) オブジェクト。 

## <a name="see-also"></a>関連項目

- [カスタム関数のリファレンスドキュメント](/javascript/api/custom-functions-runtime)
- [Excel JavaScript API の要件セット](../reference/requirement-sets/excel-api-requirement-sets.md)
