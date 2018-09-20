# <a name="metadata-element"></a>メタデータ要素

Excel でユーザー定義関数によって使用されるメタデータの設定を定義します。

## <a name="attributes"></a>属性

なし

## <a name="child-elements"></a>子要素

|  要素  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  はい  | カスタム関数で使用される JSON ファイルのリソース id を持つ文字列です。 |

## <a name="example"></a>例

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
