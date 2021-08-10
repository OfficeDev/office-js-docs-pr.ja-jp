---
title: マニフェスト ファイル内のイベント要素
description: アドインでイベント ハンドラーを定義します。
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 486236f2c2dc19f835e06bad027b4fca33809fb257ba6f6d455add66ab5b5ce0
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093298"
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
