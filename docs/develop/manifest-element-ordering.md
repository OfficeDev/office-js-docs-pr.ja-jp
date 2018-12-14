---
title: マニフェスト要素の正しい順序を確認する方法
description: 親要素内で子要素を配置するための正しい順序を確認する方法について説明します。
ms.date: 11/16/2018
ms.openlocfilehash: 3efc95926b7562b0e68bbb6f4b13c47cc4ae6824
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270615"
---
# <a name="how-to-find-the-proper-order-of-manifest-elements"></a>マニフェスト要素の正しい順序を確認する方法

Office アドインのマニフェストの XML 要素は適切な親要素の下に配置する必要があり、*また*、親要素の下で子要素同士が特定の順序に配置する必要があります。

必要な順序は、[[スキーマ](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)] フォルダー内の XSD ファイルで指定されています。 XSD ファイルは、作業ウィンドウ、コンテンツ、およびメール アドインのサブフォルダーに分類されます。

例えば、`<OfficeApp>` 要素では、`<Id>`、`<Version>`、`<ProviderName>` はこの順序で表示する必要があります。 `<AlternateId>` 要素が追加された場合、この要素は `<Id>` 要素と `<Version>` 要素の間に配置する必要があります。 順序が間違っている要素が 1 つでもあると、マニフェストは有効にならず、アドインも読み込まれません。

> [!NOTE]
> [Office アドイン検証ツール](/office/dev/add-ins/testing/troubleshoot-manifest#validate-your-manifest-with-the-office-add-in-validator)では、要素の順序が間違っている場合と、要素が間違った親の下にある場合とで、同じエラー メッセージが使用されます。 エラーには、子要素が親要素の有効な子ではないと表示されます。 そのようなエラーが表示されるものの、子要素のレファレンス ドキュメントがこの子要素は親要素の有効な子*である*と示す場合は、おそらく、子要素が間違った順序で配置されていることが原因です。

ある親要素の子要素の正しい順序を確認するには、次の手順を実行します。 (XDSD ファイルは非常に複雑なため、ここでご紹介するのは簡素化されたプロセスです。 XSD ファイルの詳細な説明はこのドキュメントの目的の範囲外です。)

1. 作成するアドインの種類の [[スキーマ](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)] の下にある サブフォルダーを開きます。 
2. 親要素が複合型として定義されている XSD ファイルを開きます。 どのファイルが複合型と定義されているかがわからない場合は、確認できるまで複数のファイルで手順 3 を行う必要があります。
3. `<xs:complexType name="PARENT_ELEMENT">` を検索します。"PARENT_ELEMENT" は親要素の名前です。
4. PARENT_ELEMENT の定義の中に、`<xs:sequence>` という名前の要素が (通常は) あります。 [TaskPaneAppVersionOverridesV1_0.xsd](https://raw.githubusercontent.com/OfficeDev/office-js-docs-pr/master/docs/overview/schemas/taskpane/TaskPaneAppVersionOverridesV1_0.xsd) での `<SuperTip>` の定義を次に示します。

```xml
  <xs:complexType name="Supertip">
    <xs:annotation>
      <xs:documentation>
        Specifies the super tip for this control.
      </xs:documentation>
    </xs:annotation>
    <xs:sequence>
      <xs:element name="Title" type="bt:ShortResourceReference" minOccurs="1" maxOccurs="1" />
      <xs:element name="Description" type="bt:LongResourceReference" minOccurs="1" maxOccurs="1" />
    </xs:sequence>
  </xs:complexType>
```

`<xs:sequence>` は、含まれる可能性のある子要素を*正しい表示順序で*一覧表示します。 ここに含まれる要素が必須という意味では*ありません*。 子要素の `minOccurs` の値が **0** の場合、この子要素は省略可能です。 *ただし、この子要素が含まれる場合は、`<xs:sequence>` 要素で指定された順序で配置する必要があります*。

`<xs:sequence>` 要素がない場合、または*ある*場合でもこの子要素が一覧表示されていない場合 (子要素のレファレンス ドキュメントで、この子要素が親要素の有効な子で*ある*と示される場合でも)、XSD ファイルのどこかで親要素の複合型の定義が追加の子要素によって拡張されています。 たとえば、`OfficeApp` 複合型の定義は、使用可能な子として `Requirements` を一覧に表示しません。 ただし、ファイルの後の方で (`TaskPaneApp` 複合型 の定義の中で) `OfficeApp` の定義は拡張され、追加の有効な子として `Requirements` が追加されています。

拡張された定義を探すには次の手順を実行します。

1. ファイルの最初から初めて、`<xs:extension base="PARENT_ELEMENT">` を検索します。"PARENT_ELEMENT" は親要素の名前です。 拡張された定義は複数ある可能性があります。
2. 作業中のコンテキストに関連する拡張された定義を探します。 たとえば、`OfficeApp` 複合型は `ContentApp`、`MailApp` および `TaskPaneApp` 複合型の中でそれぞれ拡張されます。

ファイル内の各 `<xs:extension base="PARENT_ELEMENT">` には独自の `<xs:sequence>` があり、親要素に対して有効な追加の子要素を一覧表示します。 拡張された一覧上の子要素は、常に親要素の複合型定義内の元のリストの子要素の*後ろに*配置する必要があります。

## <a name="see-also"></a>関連項目

- [Office アドイン マニフェストのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md)
