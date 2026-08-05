---
title: 数据库字段加解密与搜索方案
date: 2026-08-05 09:00:00 +0800
categories: [Blogging, Security]
tags: [security]
description: 数据库敏感字段的应用层加解密、精确与模糊搜索、数据迁移及密钥管理方案。
---

# 数据库加密字段

数据库 MySQL 和 MongoDB。需要加密的字段有：身份证，手机号码。日志不要打印出敏感数据。

## 方案约束

加密数据只使用一个字段保存，并支持直接精确查询，所以相同明文在相同密钥下会生成相同密文。这样会暴露重复数据和出现次数，是为了搜索做的折中。

能解决数据库文件、备份和数据库账号泄露后直接看到明文的问题，前提是密钥没有一起泄露。

不能解决应用服务器或密钥已经被入侵、业务接口越权，以及日志、缓存、消息队列里已经出现明文的问题。

模糊查询需要额外的分段映射表，会增加存储量，查询结果也需要解密后再过滤。

# 实操层面

首选方案

加密字段数据只能写JAVA 代码使用AES(具体是 AES/GCM/NoPadding) 来加密, 然后落库到数据库.  优点是快速, 精确. 缺点是不做封装的话代码会比较多. 如果是操作对象, 把加密和解密写进getter和setter就行.

> 为了使相同明文生成相同密文，便于使用一个字段进行精确查询，本方案不使用随机 IV。这不是 AES-GCM 的标准安全用法，是为了满足搜索需求做的折中。

使用命令启动项目时通过环境变量或启动参数传入进去, 初始化AESCryptoCipher需要拿已知的密码进行校验, 输错Secret Key对数据进行加错密

如果是用K8S则先把值配置到secret里面, 再通过环境变量或启动参数传递进去

其他方案

MyBatis 拦截器自动加解密, 配置哪些表和字段需要加解密. 优点是全自动, 缺点局限性高复杂度高兼容性低; 遇到join表无法得知重名字段是来源于哪个表, 那就需要引入SQL解析增加复杂度. mongodb 同理

# 搜索

如果是精准搜索的话，直接把搜索的值加密后再在sql里精准匹配

数据被加密后是没有办法通过sql 语句进行模糊查询. 下面方案是实现成本排序

方案一: 利用数据库的加解密, 优点是非常方便, 缺点性能极低, 需要数据库全表拉出来解密后再搜索

方案二: 密文字段存储 前N位明文 + 后几位加密, 比如13425\*\*\*\*7078,  优点实现成本低基本够用; 缺点出现部分泄露和模糊长度受限

方案三: 增加字段存储 前N位明文 + 后几位脱敏, 比如 13425783ebg6X==,  优点实现成本低基本够用; 缺点增加存储字段有可能和模糊长度受限

方案四: 按固定长度循环拆分, 加密落库到一张表, 每条数据都指向源数据的id. 搜索时将关键字按固定循环拆分并加密, 再在这张搜索表中进行匹配查询得出引用的源数据. 好处能全模糊; 缺点是映射表数据量会按分片数量成倍增加, 数据越多性能更差, 返回的数据会不精确需要二次筛选

### 方案四例子

数据新建表

```sql
CREATE TABLE `encrypt_value_mapping` (
  `id` int NOT NULL AUTO_INCREMENT COMMENT '自增id',
  `ref_id` int NOT NULL COMMENT '关联系统编号',
  `type` varchar(64) NOT NULL COMMENT '数据类型, 身份证/手机号码/银行卡',
  `encrypt_value` varchar(255) NOT NULL COMMENT '加密后的字符串'
) ENGINE=InnoDB  CHARSET=utf8mb4 COLLATE=utf8mb4_bin COMMENT='分段加密映射表'
```

比如将手机号码13425787078按固定长度拆分. 比如长度是3, 则拆分成 \[134,342,425,257,678,787,870,708\], 再对数据每个item进行加密并落库到表 encrypt\_value\_mapping

搜索 1342 的时候, 按固定长度3先拆分成数组\[134,342\], 将数据进行加密\[vLn4vFXK,cl1kDH\]

再用加密数据到数据库查找

```sql
select telephone from user
where id in (
select ref_id from encrypt_value_mapping where type = '手机号码' and encrypt_value in ('vLn4vFXK', 'cl1kDH')
)
```

获取到user的数据后, 在代码里解密进行二次筛选

```java
String keyword = "1342";
List<User> user = userDao.search(keyword);

//解密
for (User u : user) {
    u.setTelephone(AESCryptoCipher.getInstance().decryptFromBase64(u.getTelephone()));
}

// 二次过滤
user = user.stream().filter(i-> i.getTelephone().contains(keyword)).toList();
```

# 数据迁移

数据库表增加临时字段tmp, 使用java代码对源数据进行加密保存到字段tmp里, 如果需要搜索的话则也要处理.

数据全部处理完之后. mysql的话就删除原字段后tmp字段改为原字段名字. mongodb的话需要命令删除原字段

# 密钥(Secret Key)安全管理

由于 AES 的Secret Key需要安全保管, 泄露Secret Key会导致数据泄露风险.  保管Secret Key按安全级别从低到高

1.  把密钥作为spring boot 启动参数或环境变量传入, 密钥保存在硬盘的启动文件上, 文件限制用户读取. 被人偷走硬盘也能拿到密钥
    
2.  ~~把密钥作为spring boot 启动参数或环境变量传入, 远程SSH上去执行带密钥的启动命令. 好处密钥不存储到服务器上, 缺点是需要人手介入~~
    
3.  基于1, 硬盘增加 LUKS 加密.  但每次服务器开机需要人工远程输入硬盘密码
    
4.  应用部署在K8S里, 密钥是在secret里.  依赖于K8S自带的加密, 只要K8S的自身的主密钥不泄露, 则无法获取到密钥
    
5.  使用硬件HSM进行保存Secret Key, 但服务器重启HSM需要人工现场干预
    

其他方案, 本质都是一样, 人工干预初始化硬件, 再安全读取主密钥

1.  HSM + hashicorp vault + K8S
    
2.  HSM + hashicorp vault + Transit Auto-unseal
