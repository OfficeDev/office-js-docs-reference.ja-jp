# <a name="group-element"></a>Group 要素

タブには、UI コントロールのグループを定義します。カスタム タブで、アドインは最大 10 個のグループを作成できます。各グループのコントロールは、コントロールが表示されるタブに関係なく、6 個に制限されています。アドインは、カスタム タブ 1 つに制限されています。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  はい  | グループの一意の ID。|

### <a name="id-attribute"></a>id 属性

必須。グループの一意識別子。最大 125 文字の文字列です。マニフェスト内で一意にする必要があります。一意ではない場合、レンダリングに失敗します。

## <a name="child-elements"></a>子要素
|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Label](#label)      | はい |  CustomTab またはグループのラベル。  |
|  [Control](#control)    | はい |  1 つ以上のコントロール オブジェクトのコレクション。  |

### <a name="label"></a>Label 

必ず指定します。グループのラベルです。 **resid** 属性には、 **Resources** 要素の **ShortStrings** 要素にある **String** 要素の [id](resources.md) 属性の値を設定する必要があります。

### <a name="control"></a>Control
1 つのグループに少なくとも 1 つのコントロールが必要です。

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```