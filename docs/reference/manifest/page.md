# <a name="page-element"></a>Page 要素

Excel でカスタム関数によって使用される HTML ページの設定を定義します。

## <a name="attributes"></a>属性

なし

## <a name="child-elements"></a>子要素

|  要素  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  はい  | カスタム関数によって使用される HTML ファイルのリソース ID を持つ文字列。 |

## <a name="example"></a>例

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
