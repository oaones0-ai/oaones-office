# 部署手册

```text
版权声明：本项目已开源，但版权属于个人所有，禁止转卖。如有问题请联系https://www.oaones.com/contract%20me
```

## 1. 前言

```text
本文介绍泛微OA（E9版本）使用公开预览，更多介绍请看https://www.oaones.com/archives/weaverfilepreview
```

## 2. 前后端架构

### 2.1 前端文件

```text
前端文件集中在 `ecology/oaones` 下，使用jsp进行页面渲染。
最新版本已独立为https://github.com/oaones0-ai/oaones-scripts.git，请先下载前端组件。
```

### 2.2 后端文件

```text
后端文件集中在 `ecology/src or classbean/com/customization/oaones/` 下，使用ssm架构，通过mybatis的xml进行数据库交互。
```

### 2.3 其他文件

```text
其他为配置文件，用于泛微标准安全协议等的通行等。
项目中所有文件理论上应全部为新增文件，需要修改的文件在下方单独列出。如覆盖时提示冲突，说明与原有项目文件冲突，不可使用本开源项目。需要定制二开。
```

## 3. 功能说明

```link
https://www.oaones.com/archives/weaverfilepreview
```

## 4. 使用文档

### 4.1 文件覆盖

```text
将文件下载后，直接覆盖到 `ecology` 目录，重启resin。
```

### 4.2 配置oa

```text
logjint.properties最后增加内容，不然日志报错：
log4j.logger.mylog=INFO,mylog
log4j.appender.mylog=org.apache.log4j.DailyRollingFileAppender
log4j.appender.mylog.DatePattern='_'yyyyMMdd'.log'
log4j.appender.mylog.File=@mylog/mylog.log
log4j.appender.mylog.layout=org.apache.log4j.PatternLayout
log4j.appender.mylog.layout.ConversionPattern=%d{yyyy-MM-dd HH:mm:ss,SSS} %-5p [Thread:%t] %m%n
log4j.additivity.mylog=false
```
### 4.3 配置WEB-INF/resin-web-post.xml
没有则新增文件，仅限resin中间件，其他配置请咨询AI

### 4.4 菜单URL设置

```text
1.管理页面：/oaones\officeunit\page/admin.jsp
```
