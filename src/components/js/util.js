//在后续的wps版本中，wps的所有枚举值都会通过wps.Enum对象来自动支持，现阶段先人工定义
var WPS_Enum = {
    msoCTPDockPositionLeft:0,
    msoCTPDockPositionRight:2
}

function GetUrlPath() {
  // 获取完整URL
  const fullUrl = window.location.href;
  
  // 获取URL的路径部分（不包含协议、主机名和端口）
  const pathname = window.location.pathname;
  
  // 获取应用的根路径（包含协议、主机名、端口和应用路径）
  const rootPath = fullUrl.substring(0, fullUrl.indexOf(pathname) + pathname.lastIndexOf('/') + 1);
  
  return rootPath;
}

  function GetRouterHash() {
    if (window.location.protocol === 'file:') {
      return '';
    }
  
    return '/#'
  }

export default{
    WPS_Enum,
    GetUrlPath,
    GetRouterHash
}