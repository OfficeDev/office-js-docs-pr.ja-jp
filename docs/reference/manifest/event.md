---
title: マニフェスト ファイル内のイベント要素
description: アドインでイベント ハンドラーを定義します。
ms.date: 05/15/2020
ms.localizationpriority: medium
ms.openlocfilehash: d5ccddc64ffecd9ebc06b28eb37c0aee46dcc2f4
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153041"
---
# <a name="event-element"></a>Event 要素

アドインでイベント ハンドラーを定義します。

> [!NOTE]
> サポートと使用方法の詳細については[、「On-send feature for Outlook」を参照してください](../../outlook/outlook-on-send-addins.md)。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Type](#type-attribute)  |  はい  | 処理するイベントを指定します。 |
|  [FunctionExecution](#functionexecution-attribute)  |  はい  | イベント ハンドラーの実行スタイル (非同期または同期) を指定します。現在サポートされているのは同期イベント ハンドラーのみです。 |
|  [FunctionName](#functionname-attribute)  |  はい  | イベント ハンドラーの関数名を指定します。 |

### <a name="type-attribute"></a>Type 属性

必須です。イベント ハンドラーを呼び出すイベントを指定します。この属性の使用可能な値は、次の表のとおりです。

|  イベントの種類  |  説明  |
|:-----|:-----|
|  `ItemSend`  |  ユーザーがメッセージまたは会議出席依頼を送信すると、イベント ハンドラーが呼び出されます。  |

### <a name="functionexecution-attribute"></a>FunctionExecution 属性

必須です。`synchronous` に設定する必要があります。

### <a name="functionname-attribute"></a>FunctionName 属性

必須です。イベント ハンドラーの関数名を指定します。この値は、アドインの[関数ファイル](functionfile.md)内の関数名と一致する必要があります。

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
