# Excel导入导出工具

## 1. 设计理念

这套Excel导入导出工具以模型为中心。开发人员只需要定义好导入/导出用的业务模型对象，即可完成生成导入模板以及导入导出的实际功能，不需要再手工编写Excel相关的POI工具类代码。

相比目前其他的开源Excel导入导出工具，包含下拉选项的支持（可以自己开发下拉选项提供类来集成，还支持级联），校验功能的支持（使用hibernate-validator），以及导入时的错误信息（单元格内容解析错误的，hibernate-validator校验错误的，以及业务校验错误自定义产生的信息）再次导出成Excel（错误字段标红，最后一列增加红字错误提示等等）功能等能力。

## 2. 使用案例

具体例子请见附带的测试用例类。

### 2.1 编写模型类

要使用Excel导入导出组件，首先要定义一个或多个Excel导入导出模型类。每个模型类对应一个Sheet中的数据内容：

```java
package com.github.mrpanyu.excel;

import java.util.Date;

import org.hibernate.validator.constraints.NotBlank;

@ExcelSheet(name = "用户信息")
@SuppressWarnings("serial")
public class DemoUserExcelModel extends ExcelModelBase {

    @ExcelColumn(name = "用户工号", notes = "不同用户的工号不能重复", width = 20)
    @NotBlank(message = "用户工号不能为空")
    private String userCode;
    @ExcelColumn(name = "用户姓名", width = 10)
    @NotBlank(message = "用户姓名不能为空")
    private String userName;
    @ExcelColumn(name = "身份证号", width = 20)
    @NotBlank(message = "身份证号不能为空")
    private String idcardNo;
    @ExcelColumn(name = "手机号", width = 15)
    @NotBlank(message = "手机号不能为空")
    private String mobile;
    @ExcelColumn(name = "年龄", width = 10)
    private int age;
    @ExcelColumn(name = "出生日期", notes = "用户出生日期，请使用yyyy-MM-dd格式", width = 15, dateFormat = "yyyy-MM-dd")
    private Date birthday;
    @ExcelColumn(name = "性别", width = 6, selectionProvider = DemoExcelColumnSelectionProvider.class, selectionType = "gender")
    private String gender;
    @ExcelColumn(name = "职业类型", width = 10, selectionProvider = DemoExcelColumnSelectionProvider.class, selectionType = "jobType")
    private String jobType;
    @ExcelColumn(name = "省份", width = 10, selectionProvider = DemoExcelColumnSelectionProvider.class, selectionType = "province")
    private String homeProvince;
    @ExcelColumn(name = "城市", width = 10, selectionProvider = DemoExcelColumnSelectionProvider.class, selectionType = "city", selectionRefField = "homeProvince")
    private String homeCity;

    // ...... get/set方法等 ......

}
```

Excel模型类的相关说明：

1. Excel模型类一般需要继承ExcelModelBase类。不继承的情况只能用于简单的导出数据用，一般不建议。
2. @ExcelSheet标注上，声明对应的Sheet页的名称。
3. 每个涉及导入导出的属性上，需要标注@ExcelColumn。注意Excel模型类可以包含未标注@ExcelColumn的属性，但这些属性不参与框架层的导入/导出动作。
4. Excel导入导出工具本身是通过直接给属性取值赋值的方式来完成操作的，不受get/set方法影响。但为方便业务操作一般还是需要有get/set方法的，也可以通过lombok的@Data标注来替代。
5. 属性上也可以添加hibernate-validator提供的如@NotBlank之类的校验标注，导入时工具会自动做校验。
6. @ExcelColumn标注包含若干属性：
    - **name**: 列头上显示的名称
    - **width**: 列宽，单位大致是按一个字母或数字的宽度为1的方式计算
    - **notes**: 如果有值，导入模板及导出的Excel中，列头上会有个隐藏的批注信息
    - **dateFormat**: 一般用于Date或String类型的属性上，用于指定该列数据单元格的日期格式
    - **selectionProvider**: 用于Excel中下拉框显示，指定一个实现了ExcelColumnSelectionProvider接口的类名，由这个类提供下拉框数据
    - **selectionType**: 传给selectionProvider类的一个附加参数，这样可以避免每种下拉框都需要单独实现一个类，可以在类里面通过判断这个参数来if/else（当然主要是结合一般字典表的结构而言，可以用来表示“字典类型”字段，这样用一个selectionProvider完成所有字典表下拉框数据提供的功能）
    - **selectionRefField**: 当有级联下拉的情况（比如例子种的省市级联），用于指定级联选择上级的属性，比如上述例子种“市”属性指定的就是“省”属性。级联查询的情况，selectionProvider类返回的下拉列表中是需要包含上级的代码/名称的。
    - **horizontalAlign**: 列的水平对齐方式。

ExcelModelBase类包含如下主要方法：

- **hasError**: 整个对象是否有错误，包括导入解析时的错误以及自定义的业务错误
- **hasFieldError**: 某个属性是否有错误，包括导入解析时的错误以及自定义的业务错误
- **addFieldError**: 标识某个属性有错误及具体错误信息
- **addOtherError**: 增加一条不针对任何属性的错误，一般是业务性错误，比如两个属性不能同时为空等
- **getAllErrors**: 获取所有错误信息，包括属性错误和其他错误
- **getOriginalValue**: 获取导入的时候Excel单元格的原始值，一般是String类型，注意有可能是null（比如单元格不存在的情况）

### 2.2 编写selectionProvider类

如果导入导出功能中需要使用下拉列表，则需要提供selectionProvider类：

```java
package com.github.mrpanyu.excel;

import java.util.ArrayList;
import java.util.List;

/**
 * 示例用的一个下拉框选项提供类。实际使用过程中一般会从数据库等获取下拉选项。
 */
public class DemoExcelColumnSelectionProvider implements ExcelColumnSelectionProvider {

    @Override
    public List<ExcelColumnSelectionItem> selectionItems(String type) {
        List<ExcelColumnSelectionItem> items = new ArrayList<ExcelColumnSelectionItem>();
        if ("gender".equals(type)) {
            items.add(new ExcelColumnSelectionItem("0", "女"));
            items.add(new ExcelColumnSelectionItem("1", "男"));
        } else if ("jobType".equals(type)) {
            items.add(new ExcelColumnSelectionItem("01", "管理人员"));
            items.add(new ExcelColumnSelectionItem("02", "现场工人"));
            items.add(new ExcelColumnSelectionItem("03", "后勤人员"));
            items.add(new ExcelColumnSelectionItem("99", "其他人员"));
        } else if ("province".equals(type)) {
            items.add(new ExcelColumnSelectionItem("110000", "北京市"));
            items.add(new ExcelColumnSelectionItem("120000", "天津市"));
            items.add(new ExcelColumnSelectionItem("130000", "河北省"));
            items.add(new ExcelColumnSelectionItem("140000", "山西省"));
        } else if ("city".equals(type)) { // 城市要与省份联动，因此要包含引用的省份信息
            items.add(new ExcelColumnSelectionItem("110100", "北京市", "110000", "北京市"));
            items.add(new ExcelColumnSelectionItem("120100", "天津市", "120000", "天津市"));
            items.add(new ExcelColumnSelectionItem("130100", "石家庄市", "130000", "河北省"));
            items.add(new ExcelColumnSelectionItem("130200", "唐山市", "130000", "河北省"));
            items.add(new ExcelColumnSelectionItem("130300", "秦皇岛市", "130000", "河北省"));
            items.add(new ExcelColumnSelectionItem("140100", "太原市", "140000", "山西省"));
            items.add(new ExcelColumnSelectionItem("140200", "大同市", "140000", "山西省"));
            items.add(new ExcelColumnSelectionItem("140300", "阳泉市", "140000", "山西省"));
        }
        return items;
    }

}
```

### 2.3 生成导入模板

有了模型类以后，生成导入模板的操作非常简单：

```java
byte[] data = ExcelImportExportTools.impTemplate(DemoUserExcelModel.class, DemoUserExperienceExcelModel.class);
```

以上代码中，传入了两个模型类，即生成的模板有两个Sheet页，第一个对应DemoUserExcelModel对象，第二个对应DemoUserExperienceExcelModel对象。

返回的data，就是Excel导入模板文件内容，之后将其写入文件即可。

### 2.4 进行导入

用户填写了导入模板数据后，将文件上传到服务器进行导入。获取到导入的文件数据后，调用这个方法就可以获取到对象：

```java
List<List<Object>> models = ExcelImportExportTools.imp(in, DemoUserExcelModel.class, DemoUserExperienceExcelModel.class);
```

方法第一个入参是InputStream类型或byte[]类型，表示Excel文件的内容。

方法返回的List\<List\<Object>>对象，外层的List中每个元素表示一个Sheet页的数据，每个Sheet页的数据本身也是一个List，里面每个元素的类型，分别是入参中后面传入的几个模型类类型。

即上述方法调用返回的data对象，外层List应该有两个元素，分别对应第一第二个Sheet页，第一个子List内的元素是DemoUserExcelModel类型的，第二个子List内的元素是DemoUserExperienceExcelModel类型的。

对模型中定义了selectionProvider的属性，导入后模型对象中该属性的值是实际值，而不是显示名称。但注意因为Excel文件当中只能存储显示名称，因此实际上获取的值是从显示名称反向翻译回去的。当存在多个值有相同的显示名称的时候，会解析成第一个值。这在级联下拉的情况中可能存在，比如省市联动，北京和上海都有“市辖区”，当选项是“市辖区”的时候就会被解析成北京的市辖区，需要业务程序自行根据省的代码进行处理。

导入过程当中，可能发生各种填写错误问题，比如应该是数字格式的填写了文本，或是下拉框取值不在范围之内等等。在导入过程中出现的错误，都会通过模型类的`addFieldError`方法自动加入进去。因此在做后续的操作过程中，应该先通过`hasError`方法判断一下该模型数据是否正常。

一般来说，获取到对象后，业务程序应该先进行业务校验，业务校验通过后再进行实际数据导入操作。业务校验发现的问题，也可以通过`addFieldError`等方法回写到模型对象中供错误文件导出。

### 2.5 错误文件导出

导入时有任何对象有错误，应该给与用户一个反馈信息。一般来说是以提供一个错误Excel文件的形式来展现。当模型对象中已经包含了错误信息的时候，这个错误文件可以直接使用工具的导出方法生成：

```java
byte[] errorExportData = ExcelImportExportTools.exp(models, DemoUserExcelModel.class, DemoUserExperienceExcelModel.class);
```

以上方法中的data，格式同导入时返回的data格式。如果错误文件中也想保留非错误的记录，则可以直接用导入时返回的data对象进行导出，如果只需要包含错误的记录，则可以重新创建一个同样结构的List，只将有错误(`hasError() == true`)的数据添加进去进行导出。

导出后的文件，会根据模型类是否有错误，对响应的行进行标色，如果有明确属性错误的，对应的单元格也会进行标色，另外最后一列中会显示这条数据的一个汇总的错误信息，供用户检查。

导出的文件本身还是符合导入模板格式的，因此一般用户可以直接在这上面修正数据后再次导入。

### 2.6 一般的数据导出

常规的数据导出，和上述的错误文件导出功能其实完全一样，只是导出的模型对象不包含错误信息，因此不会有特别的标色/提示信息而已。
