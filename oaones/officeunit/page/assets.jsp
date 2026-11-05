<%--
  OfficeUnit 静态资源版本片段。

  本文件只依赖 OfficeUnit 项目自身资源，以及 /oaones/scripts 下的公共组件，
  不依赖其他一级项目，可随 OfficeUnit 功能单独部署。

  用法：在页面中静态引入本文件，并使用 aiAssetV：
      <%@ include file="/oaones/officeunit/page/assets.jsp" %>
--%>
<%
    long aiAssetV = 0L;
    String[] officeUnitAssetPaths = {
        // 公共组件
        "/oaones/scripts/js/tree.js",
        // OfficeUnit 项目本地资源
        "/oaones/officeunit/css/admin.css",
        "/oaones/officeunit/js/admin.js"
    };
    try {
        for (String officeUnitAssetPath : officeUnitAssetPaths) {
            java.net.URL officeUnitAssetUrl = application.getResource(officeUnitAssetPath);
            if (officeUnitAssetUrl != null) {
                aiAssetV = Math.max(
                        aiAssetV, officeUnitAssetUrl.openConnection().getLastModified());
            }
        }
    } catch (Exception officeUnitAssetIgnore) {
        aiAssetV = 0L;
    }
    if (aiAssetV <= 0L) {
        aiAssetV = System.currentTimeMillis();
    }
%>
