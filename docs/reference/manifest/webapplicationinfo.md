---
title: マニフェスト ファイルの WebApplicationInfo 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2ab06b7ec21bccf13039badcc94b9de0ce7b8600
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870276"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="2761a-102">WebApplicationInfo 要素</span><span class="sxs-lookup"><span data-stu-id="2761a-102">WebApplicationInfo element</span></span>

<span data-ttu-id="2761a-103">Office アドインでシングル サインオン (SSO) をサポートします。この要素には、次の両方としてのアドインに関する情報が含まれます。</span><span class="sxs-lookup"><span data-stu-id="2761a-103">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="2761a-104">Office ホスト アプリケーションでアクセス許可を必要とする可能性のある対象の OAuth 2.0 *リソース*。</span><span class="sxs-lookup"><span data-stu-id="2761a-104">An OAuth 2.0 *resource* to which the Office host application might need permissions.</span></span>
- <span data-ttu-id="2761a-105">Microsoft Graph に対するアクセス許可を必要とする可能性のある OAuth 2.0 *クライアント*。</span><span class="sxs-lookup"><span data-stu-id="2761a-105">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="2761a-106">現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="2761a-106">The single sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="2761a-107">シングル サインオン API の現在のサポート状態に関する詳細は、「[Identity API の要件セット](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2761a-107">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span> <span data-ttu-id="2761a-108">Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="2761a-108">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="2761a-109">この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2761a-109">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="2761a-110">**WebApplicationInfo** は、マニフェスト内の [VersionOverrides](versionoverrides.md) 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="2761a-110">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="2761a-111">子要素</span><span class="sxs-lookup"><span data-stu-id="2761a-111">Child elements</span></span>

|  <span data-ttu-id="2761a-112">要素</span><span class="sxs-lookup"><span data-stu-id="2761a-112">Element</span></span> |  <span data-ttu-id="2761a-113">必須</span><span class="sxs-lookup"><span data-stu-id="2761a-113">Required</span></span>  |  <span data-ttu-id="2761a-114">説明</span><span class="sxs-lookup"><span data-stu-id="2761a-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2761a-115">**Id**</span><span class="sxs-lookup"><span data-stu-id="2761a-115">**Id**</span></span>    |  <span data-ttu-id="2761a-116">はい</span><span class="sxs-lookup"><span data-stu-id="2761a-116">Yes</span></span>   |  <span data-ttu-id="2761a-117">Azure Active Directory (Azure AD) v2.0 エンドポイントに登録された、アドインの関連サービスの**アプリケーション ID**。</span><span class="sxs-lookup"><span data-stu-id="2761a-117">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="2761a-118">**Resource**</span><span class="sxs-lookup"><span data-stu-id="2761a-118">**Resource**</span></span>  |  <span data-ttu-id="2761a-119">はい</span><span class="sxs-lookup"><span data-stu-id="2761a-119">Yes</span></span>   |  <span data-ttu-id="2761a-120">Azure Active Directory v2.0 エンドポイントに登録されたアドインの**アプリケーション ID URI** を指定します。</span><span class="sxs-lookup"><span data-stu-id="2761a-120">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="2761a-121">Scope</span><span class="sxs-lookup"><span data-stu-id="2761a-121">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="2761a-122">いいえ</span><span class="sxs-lookup"><span data-stu-id="2761a-122">No</span></span>  |  <span data-ttu-id="2761a-123">アドインが必要とする Microsoft Graph に対するアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="2761a-123">Specifies the permissions that the add-in needs to Microsoft Graph.</span></span>  |

> [!NOTE] 
> <span data-ttu-id="2761a-124">現時点では、アドインのリソースがそのホストと一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="2761a-124">Currently, it's necessary that your add-in's Resource matches its Host.</span></span> <span data-ttu-id="2761a-125">Office は、所有権が証明できない限り、アドインのトークンを要求できません。現在これを行うには、リソースの完全修飾ドメイン名でアドインをホストします。</span><span class="sxs-lookup"><span data-stu-id="2761a-125">Office will not request a Token for an add-in unless it can prove ownership, and today this is done by hosting the add-in under the Resource's fully-qualified domain name.</span></span>

## <a name="webapplicationinfo-example"></a><span data-ttu-id="2761a-126">WebApplicationInfo の例</span><span class="sxs-lookup"><span data-stu-id="2761a-126">WebApplicationInfo example</span></span>

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
