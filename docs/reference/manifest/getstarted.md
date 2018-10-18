# <a name="getstarted-element"></a>GetStarted 要素

アドインが、Word、Excel、PowerPoint、OneNote のホストにインストールされているときに表示される吹き出しで使用される情報を提供します。**GetStarted** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。

## <a name="child-elements"></a>子要素

| 要素                       | 必須 | 説明                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [タイトル](#title)               | はい      | アドインが機能を公開する場所を定義します。     |
| [説明](#description)   | はい      | JavaScript 関数を含むファイルの URL。|
| [LearnMoreUrl](#learnmoreurl) | いいえ       | アドインの詳細を説明するページの URL。   |

### <a name="title"></a>タイトル 

必須。吹き出しの一番上に使用するタイトル。**resid** 属性は [Resources](resources.md) セクションの **ShortStrings** 要素にある有効な ID を参照します。

### <a name="description"></a>説明

必須。吹き出しの説明/本文の内容。**resid** 属性は [Resources](resources.md) セクションの **LongStrings** 要素にある有効な ID を参照します。

### <a name="learnmoreurl"></a>LearnMoreUrl

必須。ユーザーがアドインの詳細を参照できるページの URL。**resid** 属性は [Resources](resources.md) セクションの **Urls** 要素にある有効な ID を参照します。

> [!NOTE]
> |||UNTRANSLATED_CONTENT_START|||**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.|||UNTRANSLATED_CONTENT_END||| これが利用可能になったときに URL がレンダリングされるよう、すべてのクライアントにこの URL を追加することをお勧めします。 

## <a name="see-also"></a>関連項目

次のコード サンプルでは、**GetStarted** 要素を使用しています。

* [テーブルとグラフの書式設定を操作するための Excel Web アドイン](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [Word アドインの JavaScript SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [PowerPoint アドインで Microsoft Graph を使用した Excel グラフの挿入](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
