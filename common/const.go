package common

const SheetName = "Sheet1"

// 横列展示的表格
var FocusField2ExcelMap = map[string]string{
	"原始文件名":  "原始文件名",
	"市级工单编号": "市级工单编号",
	"区级工单编号": "区级工单编号",
	"市派单时间":  "市派单时间",
	"问题分类":   "问题类型",
	"工单来源":   "工单来源",
	"处理期限":   "处理期限",
	"处理时限":   "处理时限",
	"标题":     "工单标题",
	"工单标题":   "工单标题",
	"工单描述":   "工单描述",
	"问题点位":   "问题点位",
	"姓名":     "反映人",
	"联系电话":   "反映人联系电话",
	"是否要求回电": "是否要求回电",
}

var FocusField2ExcelArr = []string{
	"原始文件名",
	"市级工单编号",
	"区级工单编号",
	"市派单时间",
	"问题类型",
	"工单来源",
	"处理期限",
	"处理时限",
	"是否要求回电",
	"工单标题",
	"工单描述",
	"问题点位",
	"反映人",
	"反映人联系电话",
	"满意度",
	"回复内容",
	"回复单位",
}

const DemoWordFileName = "demo.docx"
const TmpDocxDirName = "other"
const DocxDirNamePrefix = "docxDir"

type FieldItem struct {
	Ignore       bool   // 是否忽略
	KeyField     string // 键
	ReplaceField string // docx 文档中替换标识
	ColumnIndex  int    // excel 中列标识
	ColumnVal    string // 列值
	Sheet        string // 页
}

func NewExcel2WordFileMap() map[string]*FieldItem {
	tmp := map[string]*FieldItem{
		"市级工单编号": {
			KeyField:     "市级工单编号",
			ReplaceField: `CITY_LEVEL_ORDER_NUMBER`,
		},
		"区级工单编号": {
			KeyField:     "区级工单编号",
			ReplaceField: `DISTRICT_LEVEL_ORDER_NUMBER`,
		},
		"回复单位": {
			KeyField:     "回复单位",
			ReplaceField: `REPLY_GOVERNING_BODY`,
		},
		"工单标题": {
			KeyField:     "工单标题",
			ReplaceField: `ORDER_TITLE`,
		},
		"工单描述": {
			KeyField:     "工单描述",
			ReplaceField: `ORDER_DESC`,
		},
		"回复内容": {
			KeyField:     "单位回复",
			ReplaceField: `RESPONSE`,
		},
	}

	for _, v := range tmp {
		v.ColumnIndex = -1
	}
	return tmp
}
