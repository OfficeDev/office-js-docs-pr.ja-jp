# <a name="sourcelocation-element"></a>SourceLocation 要素

Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。

## <a name="attributes"></a>属性

| **属性** | **必須** | **説明**                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| resid         | はい          | マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。 |

## <a name="child-elements"></a>子要素

なし

## <a name="example"></a>例

```xml
<SourceLocation resid="pageURL"/>
```