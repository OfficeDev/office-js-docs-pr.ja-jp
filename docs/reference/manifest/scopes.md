---
title: マニフェスト ファイルの Scopes 要素
description: 範囲要素には、アドインが外部リソースに接続するために必要なアクセス許可が含まれています。
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: be68033e86de736703d9d1593ad361918d5a147d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612242"
---
# <a name="scopes-element"></a><span data-ttu-id="20dfd-103">Scopes 要素</span><span class="sxs-lookup"><span data-stu-id="20dfd-103">Scopes element</span></span>

<span data-ttu-id="20dfd-104">アドインが外部リソース (Microsoft Graph など) に対して必要とするアクセス許可が含まれます。</span><span class="sxs-lookup"><span data-stu-id="20dfd-104">Contains permissions that the add-in needs to an external resource, such as Microsoft Graph.</span></span> <span data-ttu-id="20dfd-105">Microsoft Graph がリソースの場合、AppSource はスコープ要素を使用して同意ダイアログボックスを作成します。</span><span class="sxs-lookup"><span data-stu-id="20dfd-105">When Microsoft Graph is the resource, AppSource uses the Scopes element to create a consent dialog box.</span></span> <span data-ttu-id="20dfd-106">ユーザーがストアからアドインをインストールすると、ユーザーの Microsoft Graph のデータに対する指定されたアクセス許可をアドインに付与するように要求されます。</span><span class="sxs-lookup"><span data-stu-id="20dfd-106">When users install the add-in from the Store, they are prompted to grant the add-in the specified permissions to the user's Microsoft Graph data.</span></span>

<span data-ttu-id="20dfd-107">**スコープ**は、マニフェスト内の[Webapplicationinfo](webapplicationinfo.md)要素と[Authorization](authorization.md)要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="20dfd-107">**Scopes** is a child element of the [WebApplicationInfo](webapplicationinfo.md) and [Authorization](authorization.md) elements in the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="20dfd-108">子要素</span><span class="sxs-lookup"><span data-stu-id="20dfd-108">Child elements</span></span>

|  <span data-ttu-id="20dfd-109">要素</span><span class="sxs-lookup"><span data-stu-id="20dfd-109">Element</span></span> |  <span data-ttu-id="20dfd-110">必須</span><span class="sxs-lookup"><span data-stu-id="20dfd-110">Required</span></span>  |  <span data-ttu-id="20dfd-111">説明</span><span class="sxs-lookup"><span data-stu-id="20dfd-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="20dfd-112">**Scope**</span><span class="sxs-lookup"><span data-stu-id="20dfd-112">**Scope**</span></span>                |  <span data-ttu-id="20dfd-113">はい</span><span class="sxs-lookup"><span data-stu-id="20dfd-113">Yes</span></span>     |   <span data-ttu-id="20dfd-114">アクセス許可の名前。たとえば、[すべて] または [プロファイル] を参照します。</span><span class="sxs-lookup"><span data-stu-id="20dfd-114">The name of a permission; for example, Files.Read.All or profile.</span></span> |

## <a name="example"></a><span data-ttu-id="20dfd-115">例</span><span class="sxs-lookup"><span data-stu-id="20dfd-115">Example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc<Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
