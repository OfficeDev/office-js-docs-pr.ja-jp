# <a name="event-element"></a>Event 要素

アドインでイベント ハンドラを定義します。

> [!NOTE] 
> `Event` 要素は現在、Office 365 の Outlook on the web でのみサポートされています。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [型](#type-attribute)  |  はい  | 処理するイベントを指定します。 |
|  [FunctionExecution](#functionexecution-attribute)  |  はい  | イベント ハンドラの実行スタイル (非同期または同期) を指定します。現在サポートされているのは同期イベント ハンドラのみです。 |
|  [FunctionName](#functionname-attribute)  |  はい  | イベント ハンドラの関数名を指定します。 |

### <a name="type-attribute"></a>Type 属性

必須です。イベント ハンドラを呼び出すイベントを指定します。この属性の使用可能な値は、次の表のとおりです。

|  イベントの種類  |  説明  |
|:-----|:-----|
|  `ItemSend`  |  ユーザーがメッセージまたは会議出席依頼を送信すると、イベント ハンドラが呼び出されます。  |

### <a name="functionexecution-attribute"></a>FunctionExecution 属性

必須。 `synchronous` に設定する必要があります。

### <a name="functionname-attribute"></a>FunctionName 属性

必須です。イベント ハンドラの関数名を指定します。この値は、アドインの [ 関数ファイル](functionfile.md)内の関数名と一致する必要があります。

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```