# Excel数据管理助手
在制造业的工厂实践中，使用最广的记录数据的工具是Excel电子表格，在进行智能制造升级转型中，就势必需要将历史的Excel数据，存储到数据库中，以便数据的管理和分析。但这些Excel数据可以说格式和数据千差万别，就需要有一种有效的工具来流畅的进行数据的迁移。同时，限于信息化和计算机技术掌握的水平，大量制造业工人仍旧习惯或愿意使用Excel作为日常数据的录入工具，所以在升级转型中，仍需要考虑某些不能使用智能化机器采集的数据录入问题，一种有效的方式是保留Excel格式录入，录入完成后通过简单操作就可以导入数据库中。本Excel工具就是为解决历史数据的导入和日常数据录入的目的而来。
           
## 一、Excel历史数据导入和管理
### 0、主页面加载什么显示什么
记录上次操作的上下文，打开应用后，自动回复到上次工作的上下文中，提高效率。
这需要记录工作的步骤，一个完整的工作任务的步骤，完成到第几项，下次直接恢复到第几项。

### 1、Excel文件的加载和管理
在工具中，首先要跟MS的Excel工具风格和操作习惯一致。简言之，是一个简化版本的Excel。
**原则：要对操作过的文件有详细的保存和记录**
通过文件菜单打开需要导入的文件，或者从历史导入的记录中选择文件进行修改重新导入。
需要有一个历史记录表，创建表的SQL如下：

```sql
CREATE TABLE [dbo].[t_excel_files] (
  [oper_id] int  IDENTITY(1,1) NOT NULL,
  [b_data] varbinary(max)  NOT NULL,
  [g_time] datetime  NULL,
  [operator] varchar(30) COLLATE Chinese_PRC_CI_AS  NULL,
  [file_name] varchar(255) COLLATE Chinese_PRC_CI_AS  NULL
)
```
通过工具的保存功能，可以及时的将文件保存到硬盘，同时存储在上面这张表中，以便可以从历史记录中取出来，进行修改等操作。取出的界面如下设计

![7c9034a408e3bfce73df139ae4406e41.png](en-resource://database/2148:1)


用户可以通过对列表项进行双击操作调入，也可以对有的无用的记录进行删除。

### 2、Excel数据导入的配置
我们假定 **从在工具里打开Excel文件到把里面数据导入到数据库** 这一过程叫做**一次数据导入操作**，导入配置也是针对一次数据导入操作进行，一个Excel文件可以包括一个或多个Sheet页。导入配置包括确定数据区，确定操作的Sheet页，要导入的目标数据库、对应的数据库表，对应的字段匹配以及相关的存储过程等。配置是通过文件 oper_id 为主键进行标识。主要分为两方面的配置：

* 表格化相关配置
* 数据库相关配置
* 转化适配器配置

#### 2.1 表格化配置
本工具是通过Excel的表(Table)来确定需要被操作的数据区，需要确定数据区的左上角（单元格名称）和右下角（单元格名称）区域。然后需要确定数据区应用的Sheet页，需要假定本次导入操作对 不同Sheet页具有同样的待处理数据区。如果确实需要对该数据文件进行多次导入操作，则需要定义另一次数据导入操作配置。完成了数据区域指定和Sheet页的配置后，即可以进行 标记表的程序自动化操作。界面设计如下：
![de8af06b711274923b9a55830fe78390.png](en-resource://database/2172:1)


#### 2.2 数据库配置
导入数据库时，需要知道导入哪个数据库？哪个表？以及数据区的单元表格与数据库表字段的映射。 这也可以通过配置界面完成。
![69a9f0c653242391f395a3c6ea0d55fa.png](en-resource://database/2174:1)


#### 2.3 配置转换适配器
由于实际情况千差万别，为了适应这种情况，往往不是简单的字段映射这么简单，这也是为什么很多数据导入工具不适用的原因。本工具采用可外挂转换适配器源码的方式解决，这也是本工具一个有点或特色。外挂的源代码可以在本工具进行运行时编译并运行，目前支持C#语言。编好代码后，可以通过配置将其纳入本次数据导入范围内。关于编写转换适配器源码的设计，在下面说明。

#### 2.4 数据库表设计
需要对设置进行保存，数据库表如下：
```sql
CREATE TABLE [dbo].[t_excel_settings] (
  [id] int  IDENTITY(1,1) NOT NULL,
  [oper_id] int  NOT NULL,
  [left_upper] varchar(10) COLLATE Chinese_PRC_CI_AS  NULL,
  [right_down] varchar(10) COLLATE Chinese_PRC_CI_AS  NULL,
  [is_format] varchar(2) COLLATE Chinese_PRC_CI_AS  NULL,
  [m_time] datetime  NULL,
  [g_time] datetime  NULL,
  [sheet_list] varchar(2550) COLLATE Chinese_PRC_CI_AS  NULL,
  [db_name] varchar(255) COLLATE Chinese_PRC_CI_AS  NULL,
  [db_table] varchar(255) COLLATE Chinese_PRC_CI_AS  NULL,
  [dat_convertor] varchar(255) COLLATE Chinese_PRC_CI_AS  NULL
)
CREATE TABLE [dbo].[t_excel_dbmapings] (
  [id] int  IDENTITY(1,1) NOT NULL,
  [oper_id] int  NOT NULL,
  [column_name] varchar(255) COLLATE Chinese_PRC_CI_AS  NULL,
  [excel_index] int  NULL,
  [comments] varchar(255) COLLATE Chinese_PRC_CI_AS  NULL,
  [g_time] datetime  NULL
)
```
#### 2.5 界面设计
同时要保存本次操作的所应用的Sheet页范围，界面如下：
![be0ad3c10356fd9bef00800197a4bfe1.png](en-resource://database/2176:1)



通过保存设置按钮对设置进行保存。
##### 2.5.1 界面逻辑
打开界面时，跟进文件id从数据库中读取相关配置，更新界面元素值，如果没有配置则界面初始化为默认值。
表格化可以进行的条件是配置了数据区和操作的sheet页。
数据库相关配置可以在确定好后进行配置。
本界面可以多次打开，多次配置。

![d418fb06fdd2929a10ab8103b0df4836.png](en-resource://database/2170:1)


### 3、转换适配器的设计
为了把上述格式化后得到的原始DataTable类型的数据导入数据库中对应的表，除了上述简单的字段匹配配置外，最重要的是可以根据规则编写一个转化的C#类赖完成这项工作。
这就提出了以下几个要求：
1、工具必须具有实时编译功能，能把C#类编译并执行
2、C#类编写规范，比如要实现哪些接口，要注意什么等
3、在Excel工具上还得能管理这些C#类，比如上传、浏览功能，简单修改C#类功能。
4、能够在导入设置界面上关联C#类，在以上导入设置中已经所体现。

#### 3.1 CS文件的保存
CS文件需要在工具特定目录下保存，这样便于实时编译；同时也可以把它保存到数据库中，如果需要可以从数据库中拉取出来。

#### 3.2 CS文件的管理
cs类需要完成编写，上传，存储，使用，作废这样一个过程。
cs类的编写可以使用VS Code等编辑器完成，也可以使用本工具的RichEditor进行编辑。
上传可以通过工具的打开文件另存既可以实现保存到特定位置。
##### 3.2.1 数据库表设计
存储方式跟Excel存储方式一样，本来可以放一个表，使用类型字段来区分，但是由于该表查询比较慢，所以设计另外一个同样结构的表来存储CS文件。

```sql
CREATE TABLE [dbo].[t_cs_files] (
  [oper_id] int  IDENTITY(1,1) NOT NULL,
  [b_data] varbinary(max)  NOT NULL,
  [g_time] datetime  NULL,
  [operator] varchar(30) COLLATE Chinese_PRC_CI_AS  NULL,
  [file_name] varchar(255) COLLATE Chinese_PRC_CI_AS  NULL
)
```

##### 3.2.2 典型的使用场景
通过文件打开按钮打开 编写好的CS文件，并另存到特定目录。如果需要也可以存储到数据库中。
![aac4322d77ecb1240d55ef902488a5ed.png](en-resource://database/2164:1)

#### 3.3 转换类编写规范
##### 3.3.1 构建DataTable
根据目标数据库表，创建对应的DataTable对应，然后根据业务规则填充这个DataTable,然后导入数据库。
##### 3.3.2 业务规则的实现
业务规则的实现，其实就是一个转换器，从原始DataTable变为目标的DataTable之间的转换器，实现这个转换器，通过写一个C#类的转换器来实现。先定义好这样一个类需要实现的接口，当然这是一种做法；也可以不用定义类的规则，让用户完全自我控制。
##### 3.3.3 实时编译并运行
业务规则用C#类实现后，怎么生效呢？因为我们最终交付的是一个用户使用的Excel数据管理工具，不可能让用户去重编译代码，所以以上业务规则实现后，用户需要点击界面相关操作，就可以编译以上业务代码并允许得到结果。所以，需要我们的工具能**实时编译C#代码**并运行C#代码。这里就引出两个问题，要设计友好的操作界面让用户容易操作；另一方面也要提供编上传代码文件的友好界面。

### 4、执行导库操作
#### 4.1 有效性检查
由于每列数据的类型、格式不一致，会造成数据导入数据库中会出现错误。通过工具提供的功能进行有效性检查。
并提供处有效性检查结果。

#### 4.2 执行导入数据库的操作
2、并做好日志记录（总条数，以及导入条数，错误条数等），包括当前文件名，内容，以及关联的设置(数据区指定，对应的表，以及表设置等)，以便后续查询历史导入记录和重新导入操作。


| operId |fileName  |totalRecords  |okRows  | errRows  | processTime  |processFlag  |
| --- | --- | --- | --- | --- | --- | --- |
|  p1001| 2022.xlsx | 5000  | 4000  | 1000  |2022-08-26 13:40:02  | true  |

### 关键点记录
1、Context
2、DB Log for data import
3、Data Validation Check

### 未决问题
#### 重复导入怎么办？
从导入历史数据中，选择需要重导的文件，载入修改内容进行重导入操作
通过Excel工具加载文件进行导入，加载时，搜索历史数据，如果历史数据里有，提示用户是否覆盖历史数据？

如何标识一次导入数据库操作？

#### 解决问题列表

| 序号 |问题  |方案  |
| --- | --- | --- |
| 1 | 设置RichEditControl的字体 |  |

#### 动态编译

[\[Dynamic Class Creation in C# - Preserving Type Safety in C# with Roslyn | DotNetCurry\]](https://www.dotnetcurry.com/csharp/dynamic-class-creation-roslyn) This Article gives the example and tutorial for Roslyn.


[c# - Roslyn, how can I instantiate a class in a script during runtime and invoke methods of that class? - Stack Overflow](https://stackoverflow.com/questions/47219017/roslyn-how-can-i-instantiate-a-class-in-a-script-during-runtime-and-invoke-meth) This Article fixed Roslyn compilation issue under .Net Framework environment.


[C#实现类似navicat一样操作MySQL数据库的界面（MyBatis逆向工程思路）_aaaaabin的博客-CSDN博客_仿navicat](https://blog.csdn.net/weixin_44490080/article/details/102878659)
