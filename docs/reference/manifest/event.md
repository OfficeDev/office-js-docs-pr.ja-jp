---
title: マニフェストファイルの Event 要素
description: アドインでイベント ハンドラーを定義します。
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 3d8e94c10bed214dd976b3048e11328f10f99325
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611548"
---
# <a name="event-element"></a>Event 要素

アドインでイベント ハンドラーを定義します。

> [!NOTE]
> サポートと使用法の詳細については、「 [Outlook アドインの送信時機能](../../outlook/outlook-on-send-addins.md)」を参照してください。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [種類](#type-attribute)  |  はい  | 処理するイベントを指定します。 |
|  [FunctionExecution](#functionexecution-attribute)  |  はい  | イベント ハンドラーの実行スタイル (非同期または同期) を指定します。現在サポートされているのは同期イベント ハンドラーのみです。 |
|  [FunctionName](#functionname-attribute)  |  はい  | イベント ハンドラーの関数名を指定します。 |

### <a name="type-attribute"></a>Type 属性

必須です。イベント ハンドラーを呼び出すイベントを指定します。この属性の使用可能な値は、次の表のとおりです。

|  イベントの種類  |  説明  |
|:-----|:-----|
|  `ItemSend`  |  ユーザーがメッセージまたは会議出席依頼を送信すると、イベント ハンドラーが呼び出されます。  |

### <a name="functionexecution-attribute"></a>FunctionExecution 属性

必須です。 に設定する必要があります。

### <a name="functionname-attribute"></a>FunctionName 属性

必須です。イベント ハンドラーの関数名を指定します。この値は、アドインの[関数ファイル](functionfile.md)内の関数名と一致する必要があります。

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
