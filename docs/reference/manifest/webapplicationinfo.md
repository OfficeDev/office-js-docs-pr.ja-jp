---
title: マニフェスト ファイルの WebApplicationInfo 要素
description: Office アドインのマニフェスト (XML) ファイル用の WebApplicationInfo 要素の参照ドキュメント。
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: 8644529d82204cb9fbc07c6fe9f8a35b60a512c8
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293808"
---
# <a name="webapplicationinfo-element"></a><span data-ttu-id="d4c20-103">WebApplicationInfo 要素</span><span class="sxs-lookup"><span data-stu-id="d4c20-103">WebApplicationInfo element</span></span>

<span data-ttu-id="d4c20-104">Office アドインでシングル サインオン (SSO) をサポートします。この要素には、次の両方としてのアドインに関する情報が含まれます。</span><span class="sxs-lookup"><span data-stu-id="d4c20-104">Supports single sign-on (SSO) in Office Add-ins. This element contains information for the add-in as both:</span></span>

- <span data-ttu-id="d4c20-105">Office クライアントアプリケーションがアクセス許可を必要とする可能性のある OAuth 2.0 *リソース* 。</span><span class="sxs-lookup"><span data-stu-id="d4c20-105">An OAuth 2.0 *resource* to which the Office client application might need permissions.</span></span>
- <span data-ttu-id="d4c20-106">Microsoft Graph に対するアクセス許可を必要とする可能性のある OAuth 2.0 *クライアント*。</span><span class="sxs-lookup"><span data-stu-id="d4c20-106">An OAuth 2.0 *client* that might need permissions to Microsoft Graph.</span></span>

> [!NOTE]
> <span data-ttu-id="d4c20-107">現在、シングルサインオン API は Word、Excel、Outlook、および PowerPoint でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="d4c20-107">The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="d4c20-108">シングル サインオン API の現在のサポート状態に関する詳細は、「[Identity API の要件セット](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d4c20-108">For more information about where the single sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span> <span data-ttu-id="d4c20-109">Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="d4c20-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="d4c20-110">この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d4c20-110">To learn how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="d4c20-111">**WebApplicationInfo** は、マニフェスト内の [VersionOverrides](versionoverrides.md) 要素の子要素です。</span><span class="sxs-lookup"><span data-stu-id="d4c20-111">**WebApplicationInfo** is a child element of the [VersionOverrides](versionoverrides.md) element in the manifest.</span></span>  

## <a name="child-elements"></a><span data-ttu-id="d4c20-112">子要素</span><span class="sxs-lookup"><span data-stu-id="d4c20-112">Child elements</span></span>

|  <span data-ttu-id="d4c20-113">要素</span><span class="sxs-lookup"><span data-stu-id="d4c20-113">Element</span></span> |  <span data-ttu-id="d4c20-114">必須</span><span class="sxs-lookup"><span data-stu-id="d4c20-114">Required</span></span>  |  <span data-ttu-id="d4c20-115">説明</span><span class="sxs-lookup"><span data-stu-id="d4c20-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="d4c20-116">**Id**</span><span class="sxs-lookup"><span data-stu-id="d4c20-116">**Id**</span></span>    |  <span data-ttu-id="d4c20-117">はい</span><span class="sxs-lookup"><span data-stu-id="d4c20-117">Yes</span></span>   |  <span data-ttu-id="d4c20-118">Azure Active Directory (Azure AD) v2.0 エンドポイントに登録された、アドインの関連サービスの**アプリケーション ID**。</span><span class="sxs-lookup"><span data-stu-id="d4c20-118">The **Application Id** of the add-in's associated service as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  <span data-ttu-id="d4c20-119">**MsaId**</span><span class="sxs-lookup"><span data-stu-id="d4c20-119">**MsaId**</span></span>    |  <span data-ttu-id="d4c20-120">いいえ</span><span class="sxs-lookup"><span data-stu-id="d4c20-120">No</span></span>   |  <span data-ttu-id="d4c20-121">Msm.live.com に登録されている、アドインの web アプリケーションのクライアント ID。</span><span class="sxs-lookup"><span data-stu-id="d4c20-121">The client ID of your add-in's web application for MSA as registered in msm.live.com.</span></span>|
|  <span data-ttu-id="d4c20-122">**Resource**</span><span class="sxs-lookup"><span data-stu-id="d4c20-122">**Resource**</span></span>  |  <span data-ttu-id="d4c20-123">はい</span><span class="sxs-lookup"><span data-stu-id="d4c20-123">Yes</span></span>   |  <span data-ttu-id="d4c20-124">Azure Active Directory v2.0 エンドポイントに登録されたアドインの**アプリケーション ID URI** を指定します。</span><span class="sxs-lookup"><span data-stu-id="d4c20-124">Specifies the **Application ID URI** of the add-in as registered in the Azure Active Directory v 2.0 endpoint.</span></span>|
|  [<span data-ttu-id="d4c20-125">Scope</span><span class="sxs-lookup"><span data-stu-id="d4c20-125">Scopes</span></span>](scopes.md)                |  <span data-ttu-id="d4c20-126">はい</span><span class="sxs-lookup"><span data-stu-id="d4c20-126">Yes</span></span>  |  <span data-ttu-id="d4c20-127">Microsoft Graph などのリソースに対してアドインが必要とするアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="d4c20-127">Specifies the permissions that the add-in needs to a resource, such as Microsoft Graph.</span></span>  |
|  [<span data-ttu-id="d4c20-128">Authorizations</span><span class="sxs-lookup"><span data-stu-id="d4c20-128">Authorizations</span></span>](authorizations.md)  |  <span data-ttu-id="d4c20-129">いいえ</span><span class="sxs-lookup"><span data-stu-id="d4c20-129">No</span></span>   | <span data-ttu-id="d4c20-130">アドインの web アプリケーションが承認を必要とする外部リソースと、必要なアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="d4c20-130">Specifies the external resources that the add-in's web application needs authorization to and the required permissions.</span></span>|

## <a name="webapplicationinfo-example"></a><span data-ttu-id="d4c20-131">WebApplicationInfo の例</span><span class="sxs-lookup"><span data-stu-id="d4c20-131">WebApplicationInfo example</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
