# <a name="label-component-in-office-ui-fabric"></a>Office UI Fabric のラベル コンポーネント

ラベルを使用して、コンポーネントやコンポーネントのグループに名前またはタイトルを付けます。 別のコンポーネントまたはコンポーネントのグループと組み合わせる場合、ラベルは関連するコンポーネントまたはグループのすぐ近くに配置する必要があります。 一部のコンポーネントには、ドロップダウンやトグルなど定義済みのラベルがあります。
  
#### <a name="example-label-in-a-task-pane"></a>例:作業ウィンドウ内のラベル

<br/>

![ラベルが表示された画像](../../images/overview_withApp_label.png)

<br/>

## <a name="best-practices"></a>ベスト プラクティス

|**するべきこと**|**してはいけないこと**|
|:------------|:--------------|
|たとえば **First name** など、文の先頭文字に大文字の指定を使用します。|たとえば **First Name** など、単語の先頭文字に大文字の指定を使用しないでください。|
|短くて簡潔にしてください。|完全な文章や、コロンやセミコロンなど複雑な区切り文字は使用しないでください。|
|ラベルをコンポーネントに追加する際は、ラベルのテキストに名詞または名詞の短い句を使用します。| |

## <a name="variants"></a>バリアント

|**バリエーション**|**説明**|**例**|
|:------------|:--------------|:----------|
|**既定のラベル**|標準のラベルに使用します。|![既定のラベルの画像](../../images/label.png)<br/>|
|**無効なラベル**|関連するコンポーネントが無効になっているときに使用します。|![無効なラベルの画像](../../images/labelDisabled.png)<br/>|
|**必要なラベル**|関連するコンポーネントが必要なときに使用します。|![必要なラベルの画像](../../images/labelRequired.png)<br/>|

## <a name="implementation"></a>実装

詳細については、「[ラベル](https://dev.office.com/fabric#/components/label)」と「[Fabric React のコード サンプルの使用にあたって](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact)」を参照してください。

## <a name="additional-resources"></a>その他のリソース

- [UX 設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

- [Office アドインの Office UI Fabric](office-ui-fabric.md)
