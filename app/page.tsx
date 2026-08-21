"use client";

import "antd/dist/reset.css";
import React, { useEffect, useMemo, useRef, useState } from "react";
import html2canvas from "html2canvas";
import * as XLSX from "xlsx";
import {
  App,
  AutoComplete,
  Avatar,
  Badge,
  Breadcrumb,
  Button,
  Card,
  Checkbox,
  Col,
  ConfigProvider,
  Descriptions,
  Drawer,
  Dropdown,
  Empty,
  Form,
  Input,
  InputNumber,
  Layout,
  List,
  Menu,
  message,
  Modal,
  Progress,
  Radio,
  Rate,
  Result,
  Row,
  Select,
  Space,
  Spin,
  Statistic,
  Steps,
  Table,
  Tabs,
  Tag,
  Timeline,
  Tooltip,
  Tree,
  Upload,
} from "antd";
import {
  ApartmentOutlined,
  BankOutlined,
  AppstoreOutlined,
  AuditOutlined,
  BellOutlined,
  BuildOutlined,
  CheckCircleFilled,
  CheckCircleOutlined,
  CloudUploadOutlined,
  DatabaseOutlined,
  DeploymentUnitOutlined,
  DragOutlined,
  ExperimentOutlined,
  FileDoneOutlined,
  FileSearchOutlined,
  FormOutlined,
  GlobalOutlined,
  HeartOutlined,
  HistoryOutlined,
  HomeOutlined,
  LineChartOutlined,
  LinkOutlined,
  LogoutOutlined,
  MenuFoldOutlined,
  MenuUnfoldOutlined,
  MessageOutlined,
  NodeIndexOutlined,
  NotificationOutlined,
  PlusOutlined,
  ProductOutlined,
  ProjectOutlined,
  ReloadOutlined,
  SafetyCertificateOutlined,
  SearchOutlined,
  SettingOutlined,
  SolutionOutlined,
  StarOutlined,
  SwapOutlined,
  TableOutlined,
  TeamOutlined,
  ToolOutlined,
  UserOutlined,
  ShopOutlined,
  ApiOutlined,
  RobotOutlined,
  DownloadOutlined,
} from "@ant-design/icons";
import ReactECharts from "echarts-for-react";
import {
  DndContext,
  PointerSensor,
  closestCenter,
  useSensor,
  useSensors,
  type DragEndEvent,
} from "@dnd-kit/core";
import {
  SortableContext,
  arrayMove,
  verticalListSortingStrategy,
  useSortable,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import {
  candidates,
  companyResults,
  defaultFlow,
  directoryTree,
  mappingRows,
  modules,
  products as initialProducts,
  resources,
  reviews,
  rules,
} from "./mockData";

const assetPrefix = process.env.NEXT_PUBLIC_BASE_PATH ?? "";

const { Header, Sider, Content } = Layout;
type Portal = "login" | "app" | "branch" | "manage";
type RoleKey = "business" | "branch" | "admin";
type View = string;

const roleMap = {
  business: {
    name: "业务人员",
    org: "业务应用演示角色",
    portal: "app" as Portal,
    icon: <TeamOutlined />,
  },
  branch: {
    name: "分行产品人员",
    org: "分行创新团队",
    portal: "branch" as Portal,
    icon: <BuildOutlined />,
  },
  admin: {
    name: "总行管理人员",
    org: "总行外数产品管理团队",
    portal: "manage" as Portal,
    icon: <AuditOutlined />,
  },
};

const manageMenu = [
  { key: "manage-home", icon: <HomeOutlined />, label: "管理首页" },
  { key: "manage-resources", icon: <DatabaseOutlined />, label: "外数资源管理" },
  { key: "resource-onboarding", icon: <ApiOutlined />, label: "标准化资源接入" },
  { key: "candidates", icon: <FileSearchOutlined />, label: "报送与受理" },
  { key: "build", icon: <BuildOutlined />, label: "产品建设中心" },
  { key: "assets", icon: <DeploymentUnitOutlined />, label: "能力资产中心" },
  { key: "finished", icon: <ProductOutlined />, label: "成品管理" },
  { key: "operations", icon: <LineChartOutlined />, label: "发布与运营" },
  { key: "settings", icon: <SettingOutlined />, label: "系统配置" },
];
const branchMenu = [
  { key: "branch-home", icon: <HomeOutlined />, label: "分行工作台" },
  { key: "branch-build", icon: <BuildOutlined />, label: "在线组建与试运行" },
  { key: "report", icon: <CloudUploadOutlined />, label: "在线报送" },
  { key: "branch-submissions", icon: <FileDoneOutlined />, label: "我的报送" },
  {
    key: "branch-feedback",
    icon: <MessageOutlined />,
    label: "上架与运营反馈",
  },
];
const appMenu = [
  { key: "app-home", icon: <HomeOutlined />, label: "首页" },
  { key: "app-resources", icon: <DatabaseOutlined />, label: "外数资源" },
  { key: "suppliers", icon: <ShopOutlined />, label: "数据供应商" },
  { key: "app-products", icon: <ProductOutlined />, label: "外数产品" },
  { key: "scenes", icon: <SearchOutlined />, label: "场景找数" },
  { key: "workbench", icon: <ToolOutlined />, label: "产品工作台" },
  { key: "follows", icon: <HeartOutlined />, label: "我的关注" },
  { key: "feedback", icon: <MessageOutlined />, label: "评价反馈" },
];

const reportingOrganizations = [
  "总行",
  "北京市分行",
  "天津市分行",
  "河北省分行",
  "山西省分行",
  "内蒙古自治区分行",
  "辽宁省分行",
  "吉林省分行",
  "黑龙江省分行",
  "上海市分行",
  "江苏省分行",
  "浙江省分行",
  "安徽省分行",
  "福建省分行",
  "江西省分行",
  "山东省分行",
  "河南省分行",
  "湖北省分行",
  "湖南省分行",
  "广东省分行",
  "广西壮族自治区分行",
  "海南省分行",
  "重庆市分行",
  "四川省分行",
  "贵州省分行",
  "云南省分行",
  "西藏自治区分行",
  "陕西省分行",
  "甘肃省分行",
  "青海省分行",
  "宁夏回族自治区分行",
  "新疆维吾尔自治区分行",
  "大连市分行",
  "宁波市分行",
  "厦门市分行",
  "青岛市分行",
];
const domesticRegions = [
  "全国",
  "北京市",
  "天津市",
  "河北省",
  "山西省",
  "内蒙古自治区",
  "辽宁省",
  "吉林省",
  "黑龙江省",
  "上海市",
  "江苏省",
  "浙江省",
  "安徽省",
  "福建省",
  "江西省",
  "山东省",
  "河南省",
  "湖北省",
  "湖南省",
  "广东省",
  "广西壮族自治区",
  "海南省",
  "重庆市",
  "四川省",
  "贵州省",
  "云南省",
  "西藏自治区",
  "陕西省",
  "甘肃省",
  "青海省",
  "宁夏回族自治区",
  "新疆维吾尔自治区",
  "香港特别行政区",
  "澳门特别行政区",
  "台湾省",
];
const productTaskCategories = [
  "客户营销",
  "授信支持",
  "风险监测",
  "产业研究",
  "经营分析",
  "监管与合规",
];
const demandTaskOptions: Record<string, string[]> = {
  客户营销: [
    "目标客户发现",
    "招投标线索挖掘",
    "专精特新客群筛选",
    "客户画像与分层",
    "融资机会识别",
  ],
  授信支持: [
    "客户准入核验",
    "授信调查补充",
    "关联关系识别",
    "经营稳定性核验",
    "押品与权属核验",
  ],
  风险监测: [
    "司法风险监测",
    "经营异常监测",
    "舆情风险预警",
    "供应链风险传导",
    "存量客户异动监测",
  ],
  产业研究: [
    "产业链图谱分析",
    "区域产业画像",
    "产业事件跟踪",
    "政策影响分析",
    "重点企业识别",
  ],
  经营分析: [
    "客户结构分析",
    "区域经营分析",
    "产品渗透分析",
    "机构对标分析",
    "业务机会测算",
  ],
  监管与合规: [
    "监管报送补充",
    "客户身份核验",
    "反洗钱名单核验",
    "数据合规核验",
    "关联交易识别",
  ],
};
const demandObjectOptions = [
  "企业主体",
  "集团客户",
  "个人客户",
  "行业",
  "产业链",
  "区域",
  "项目",
  "事件",
  "金融市场标的",
  "公共机构",
];
const demandFrequencyOptions = [
  "实时",
  "准实时（小时级）",
  "每日",
  "每周",
  "每月",
  "每季度",
  "事件触发",
  "按需运行",
  "一次性",
];
const defaultFollowProducts = [
  "专精特新企业营销名单",
  "企业司法风险监测",
  "区域产业链客户画像",
];

function useLiveChineseDate() {
  const formatDate = () => {
    const now = new Date();
    const weekdays = [
      "星期日",
      "星期一",
      "星期二",
      "星期三",
      "星期四",
      "星期五",
      "星期六",
    ];
    return `${now.getFullYear()}年${now.getMonth() + 1}月${now.getDate()}日 ${weekdays[now.getDay()]}`;
  };
  const [dateText, setDateText] = useState(formatDate);
  useEffect(() => {
    const timer = window.setInterval(() => setDateText(formatDate()), 1000);
    return () => window.clearInterval(timer);
  }, []);
  return dateText;
}

const buildProfiles: Record<string, any> = {
  "CP-2026-0801-09": {
    category: "地方产业企业名单筛选",
    task: "面向分行辖内半导体产业链，按地区、链属环节、企业规模、产业资质、招投标、融资事件与司法风险筛选候选企业名单",
    role: "某分行对公客户经理、产业客户团队、风险管理人员",
    process:
      "主体输入—工商清洗—链属识别—资质匹配—活跃度计算—融资事件识别—司法核验—规则筛选—名单生成",
    region: "分行辖内半导体产业链（地方产品）",
    inputs:
      "某分行半导体产业链企业名录、全国企业工商、产业链图谱、产业资质、招投标、融资事件、司法风险；行内客户关系与授信信息",
    processing:
      "主体关联、数据清洗、字段映射、指标计算、规则筛选、风险核验和名单生成",
    output: "候选企业名单、入选依据、产业链环节、基础机会信号与风险核验结果",
    action: "供客户经理查看和Excel导出；符合条件的企业进入营销或风险核验",
    flow: [
      "企业主体输入",
      "主体关联",
      "工商字段清洗",
      "半导体链属识别",
      "产业资质匹配",
      "招投标活跃度计算",
      "融资事件识别",
      "司法风险核验",
      "地方规则筛选",
      "候选名单输出",
    ],
  },
  "CP-2026-0731-01": {
    category: "客户营销",
    task: "识别浙江辖内可营销的专精特新制造业企业",
    role: "对公客户经理",
    process: "客户触达前的目标客户发现",
    region: "浙江省",
    inputs: "工商登记、专精特新名单、招投标、司法风险",
    processing: "主体关联、资质识别、活跃度计算、风险核验",
    output: "营销名单、企业画像、推荐理由与行动提示",
    action: "进入客户经理任务池，记录触达与营销结果",
    flow: [
      "企业主体输入",
      "主体关联",
      "工商信息获取",
      "专精特新资质识别",
      "招投标活跃度计算",
      "司法风险核验",
      "规则筛选",
      "名单排序",
      "营销名单输出",
      "客户经理行动提示",
    ],
  },
  "CP-2026-0728-06": {
    category: "客户营销",
    task: "识别江苏辖内近期招中标活跃且具有融资需求的企业",
    role: "对公客户经理",
    process: "营销线索发现与商机分派",
    region: "江苏省",
    inputs: "工商登记、招投标公告、行政许可、司法风险",
    processing: "主体关联、招投标聚合、融资机会评分、风险排除",
    output: "客户拓展名单、项目明细、融资机会提示",
    action: "按机构和管户关系分派客户经理跟进",
    flow: [
      "企业主体输入",
      "主体关联",
      "工商信息获取",
      "招投标事件归集",
      "项目金额与时效计算",
      "融资机会评分",
      "司法风险排除",
      "名单排序",
      "客户拓展名单输出",
      "商机任务分派",
    ],
  },
  "CP-2026-0724-03": {
    category: "风险监测",
    task: "监测广东核心企业及上下游客户的供应链风险传导",
    role: "风险管理人员",
    process: "存量客户风险监测与预警处置",
    region: "广东省",
    inputs: "供应链关系、工商变更、司法风险、负面舆情",
    processing: "关系识别、风险事件归集、传导路径计算、预警分级",
    output: "风险预警、关联图谱、重点客户处置清单",
    action: "推送风险人员核验并记录处置结论",
    flow: [
      "存量客户输入",
      "主体关联",
      "供应链关系获取",
      "工商变更监测",
      "司法风险核验",
      "负面舆情识别",
      "风险传导计算",
      "预警分级",
      "处置清单输出",
      "风险任务推送",
    ],
  },
};

const buildTabs = [
  ["agent", "受理校验"],
  ["definition", "产品定义"],
  ["scene", "场景映射"],
  ["dedupe", "查重比对"],
  ["path", "建设路径"],
  ["breakdown", "结构化拆解"],
  ["mapping", "数据契约与映射"],
  ["match", "能力/规则/模型配置"],
  ["canvas", "产品组装与流程编排"],
  ["testing", "测试验证"],
  ["simulation", "业务试点"],
  ["version", "版本定版与复制包"],
  ["publish", "审批发布"],
  ["attribution", "效果归因"],
  ["deposit", "能力回沉"],
];
const buildStages = [
  {
    name: "需求受理与定义",
    desc: "形成需求基线与产品定义书",
    tabs: ["agent", "definition", "scene"],
  },
  {
    name: "复用评估与方案",
    desc: "查重、选路并拆解产品结构",
    tabs: ["dedupe", "path", "breakdown"],
  },
  {
    name: "标准化组合建设",
    desc: "映射数据并组装母版、能力与规则",
    tabs: ["mapping", "match", "canvas"],
    focus: true,
  },
  {
    name: "联调验证与试点",
    desc: "验证数据、功能与业务闭环",
    tabs: ["testing", "simulation"],
  },
  {
    name: "定版审批与发布",
    desc: "固化复制边界并审批上架",
    tabs: ["version", "publish"],
  },
  {
    name: "运营评价与回沉",
    desc: "用运行效果驱动版本迭代",
    tabs: ["attribution", "deposit"],
  },
];

function exportExcel(
  rows: any[],
  filename = "产品结果名单.xls",
  selectedKeys?: string[],
) {
  const judicial = rows.some((r) => r.event || r.caseNo);
  const titles: Record<string, string> = {
    name: "企业名称",
    city: "地区",
    event: "事件类型",
    caseNo: "案号/文号",
    amount: "金额",
    risk: "风险等级",
    date: "日期",
    industry: "行业",
    qualification: "资质",
    bids: "招投标次数",
    score: "评分",
    project: "项目名称",
    status: "状态",
    ubo: "最终受益人",
    relations: "关联数量",
    judicial: "司法事项",
    penalty: "行政处罚",
    conclusion: "核验结论",
    source: "信息来源",
    summary: "事件摘要",
    sentiment: "情感倾向",
    chainRole: "链属环节",
    product: "核心产品",
    finance: "融资需求",
    authority: "处罚机关",
  };
  const allColumns = Object.keys(rows[0] || {})
    .filter((key) => key !== "key")
    .map((key) => ({ key, title: titles[key] || key }));
  const chosen = selectedKeys?.length
    ? allColumns.filter((x) => selectedKeys.includes(x.key))
    : allColumns;
  const columns = chosen.map((x) => x.title);
  const body = rows
    .map(
      (r) =>
        `<tr>${chosen.map((x) => `<td>${r[x.key] ?? "—"}</td>`).join("")}</tr>`,
    )
    .join("");
  const exportTime = new Date().toLocaleString("zh-CN", { hour12: false });
  const html = `<html><head><meta charset="UTF-8"><style>body{font-family:Arial,"Microsoft YaHei";position:relative;color:#222}.report-title{margin:0 0 5px;font-size:16px;line-height:1.4;font-weight:600}.meta{margin:0 0 14px;color:#777;font-size:10px}table{border-collapse:collapse;width:100%;font-size:11px}th{background:#c7000b;color:#fff;font-weight:600}th,td{padding:6px;border:1px solid #bbb}.source-mark{margin-top:10px;color:#c7000b;font-size:11px;line-height:1.5}.confidential{margin-top:5px;color:#777;font-size:10px}</style></head><body><h3 class="report-title">外部数据产品结果清单（演示）</h3><div class="meta">导出时间：${exportTime}　数据用途：内部业务参考　平台标识：ICBC-EDP-DEMO</div><table><tr>${columns.map((x) => `<th>${x}</th>`).join("")}</tr>${body}</table><p class="source-mark">资料来源：中国工商银行 外数产品管理平台</p><p class="confidential">本文件未经授权不得对外传播。</p></body></html>`;
  const url = URL.createObjectURL(
    new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" }),
  );
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
  message.success("Excel名单已导出");
}

function DemoTag() {
  return <span className="demo-tag">演示数据</span>;
}
function StatusTag({ status }: { status: string }) {
  const color =
    status.includes("正式") ||
    status.includes("通过") ||
    status.includes("正常") ||
    status.includes("已匹配")
      ? "success"
      : status.includes("待")
        ? "processing"
        : status.includes("终止") || status.includes("异常")
          ? "error"
          : "warning";
  return <Tag color={color}>{status}</Tag>;
}
function SectionTitle({
  title,
  extra,
  sub,
}: {
  title: string;
  extra?: React.ReactNode;
  sub?: string;
}) {
  return (
    <div className="section-title">
      <div>
        <h2>{title}</h2>
        {sub && <p>{sub}</p>}
      </div>
      {extra}
    </div>
  );
}
function BankBrand({ compact = false }: { compact?: boolean }) {
  return (
    <div className={"brand " + (compact ? "compact" : "")}>
      <img className="icbc-logo" src={`${assetPrefix}/icbc-logo.png`} alt="中国工商银行 ICBC" />
      <i></i>
      <strong>外部数据产品管理平台</strong>
    </div>
  );
}

function Login({ onLogin, onDemo }: { onLogin: (r: RoleKey) => void; onDemo: () => void }) {
  const [selected, setSelected] = useState<RoleKey>("business");
  return (
    <div className="login-shell">
      <header className="login-top">
        <BankBrand />
        <div className="muted">
          <GlobalOutlined /> 企业级数据管理应用 · <DemoTag />
        </div>
      </header>
      <main className="login-main">
        <div className="architecture-panel">
          <div className="headquarters-photo">
            <img src={`${assetPrefix}/icbc-headquarters.jpg`} alt="中国工商银行总行办公楼" />
            <span>中国工商银行总行办公楼</span>
          </div>
          <div className="bank-motto">
            <b>工于至诚，行以致远</b>
            <em>—— 诚信、人本、稳健、创新、卓越</em>
          </div>
        </div>
        <Card className="role-card">
          <h2>统一身份选择</h2>
          <p>请选择本次演示角色，无需输入密码</p>
          <div className="role-grid">
            {(Object.keys(roleMap) as RoleKey[]).map((k) => (
              <button
                key={k}
                className={selected === k ? "role-item active" : "role-item"}
                onClick={() => setSelected(k)}
              >
                <span className="role-icon">{roleMap[k].icon}</span>
                <b>{roleMap[k].name}</b>
                <small>{roleMap[k].org}</small>
                <em>
                  {roleMap[k].portal === "app"
                    ? "进入业务应用门户"
                    : roleMap[k].portal === "branch"
                      ? "进入分行建设与报送门户"
                      : "进入总行建设管理门户"}
                </em>
              </button>
            ))}
          </div>
          <Button
            type="primary"
            size="large"
            block
            onClick={() => onLogin(selected)}
          >
            进入平台
          </Button>
          <Button
            className="report-demo-entry"
            size="large"
            block
            icon={<DeploymentUnitOutlined />}
            onClick={onDemo}
          >
            进入汇报演示
          </Button>
          <div className="login-hint">
            <SafetyCertificateOutlined />{" "}
            本原型不连接生产数据，全部内容均为演示数据
          </div>
        </Card>
      </main>
      <footer>中国工商银行数据管理部 · 外部数据产品标准化建设机制研究</footer>
    </div>
  );
}

type DemoNode = {
  id: string;
  name: string;
  type: "input" | "common" | "new" | "output";
  function: string;
  input: string;
  output: string;
  condition: string;
  calls: string;
};

const demoBaseNodes: DemoNode[] = [
  { id: "input", name: "数据输入", type: "input", function: "汇集产品所需的内外部数据", input: "工商登记、企业资质、项目建设、招投标、融资事件、司法涉诉、经营异常", output: "待加工企业主体集", condition: "授权有效、更新时点通过校验", calls: "江苏分行半导体产业客户机会识别" },
  { id: "relation", name: "主体关联", type: "common", function: "将多源企业记录归并至行内标准主体", input: "企业名称、统一社会信用代码", output: "标准主体ID", condition: "匹配置信度不低于0.90", calls: "先进制造客户识别、新能源产业客户筛选等6项产品" },
  { id: "industry", name: "半导体链属定位", type: "new", function: "判断企业位于设计、制造、封测、设备或材料环节", input: "经营范围、产品、资质、项目与供应链信息", output: "产业链环节与判断证据", condition: "半导体产业分类口径通过评审", calls: "本次新增，支撑产业链结构与机会分布" },
  { id: "event", name: "扩产事件识别", type: "new", function: "从项目备案、环评、设备招标和厂房建设中识别真实扩产活动", input: "项目备案、设备招标、环评与建设进度", output: "扩产事件、投资规模、建设阶段、证据链", condition: "至少两类独立信号交叉验证", calls: "本次重点建设，支撑经营事件变化趋势" },
  { id: "need", name: "融资需求研判", type: "new", function: "将扩产事件转化为可核实的融资需求与产品匹配建议", input: "扩产事件、设备采购、订单与既有融资信息", output: "需求类型、时间窗口、估算区间与建议产品", condition: "仅形成营销线索，须由客户经理核实", calls: "本次重点建设，支撑潜在融资机会与预计需求" },
  { id: "risk", name: "风险核验", type: "common", function: "对入围企业进行风险排除和提示", input: "司法涉诉、经营异常、行业特定风险", output: "风险标签、证据明细与核验建议", condition: "达到重大风险阈值时转人工复核", calls: "客户营销、授信支持、风险监测等12项产品" },
  { id: "decision", name: "机会分层与行动建议", type: "common", function: "综合机会、需求和风险证据形成分层行动", input: "产业链位置、融资需求、风险标签", output: "优先访谈、方案匹配、授信核验、持续监测", condition: "结果可解释且按机构权限展示", calls: "支撑决策驾驶舱与业务任务中心" },
  { id: "output", name: "决策结果输出", type: "output", function: "形成全链路分析与业务任务", input: "产业、机会、需求、风险与行动结果", output: "产业驾驶舱、单户画像、融资建议、风险任务", condition: "按机构权限展示", calls: "产品工作台" },
];

function ReportDemo({ onExit }: { onExit: () => void }) {
  const [step, setStep] = useState(0);
  const [buildMode, setBuildMode] = useState<"analysis" | "canvas">("analysis");
  const [nodes, setNodes] = useState(demoBaseNodes);
  const [activeId, setActiveId] = useState("industry");
  const [adaptOpen, setAdaptOpen] = useState(false);
  const [adaptType, setAdaptType] = useState<"region" | "industry">("industry");
  const [adapted, setAdapted] = useState(false);
  const [running, setRunning] = useState(false);
  const [companyOpen, setCompanyOpen] = useState(false);
  const [reviewOpen, setReviewOpen] = useState(false);
  const [branchSelected, setBranchSelected] = useState(["工商登记","企业资质","招投标","司法涉诉","经营异常"]);
  const [catalogType, setCatalogType] = useState<"resource"|"asset">("resource");
  const [modulePanel, setModulePanel] = useState<"define"|"contract"|"implement"|"test">("define");
  const [moduleTested, setModuleTested] = useState(false);
  const [newModuleOpen, setNewModuleOpen] = useState(false);
  const [newModuleName, setNewModuleName] = useState("");
  const [customModules, setCustomModules] = useState<string[]>([]);
  const [dropStage, setDropStage] = useState<number | null>(null);
  const [rulesState, setRulesState] = useState({ region: "江苏省", industry: "半导体", years: 3, operation: 12, finance: 12, risk: 70, qualification: "省级以上产业资质", justice: "重大涉诉转人工复核", abnormal: "列入严重违法名单排除" });
  const active = nodes.find((n) => n.id === activeId) || nodes[0];
  const moveNode = (from: number, to: number) => {
    if (from === to || from === 0 || from === nodes.length - 1 || to === 0 || to === nodes.length - 1) return;
    const order = ["relation","industry","event","need","risk","decision"];
    const candidate = [...nodes];
    const [dragged] = candidate.splice(from, 1);
    candidate.splice(to, 0, dragged);
    const sequence = candidate.filter(n=>order.includes(n.id)).map(n=>order.indexOf(n.id));
    if (sequence.some((v,i)=>i>0&&v<sequence[i-1])) return message.warning("该位置不满足输入输出依赖：请按数据准备→产业定位→机会研判→风险校验→决策输出编排");
    const next = [...nodes];
    const [picked] = next.splice(from, 1);
    next.splice(to, 0, picked);
    setNodes(next);
  };
  const addModule = () => {
    if (nodes.some(n => n.id === "clean")) return message.info("数据清洗模块已在画布中");
    const next = [...nodes];
    next.splice(2, 0, { id: "clean", name: "数据清洗", type: "common", function: "完成格式、缺失值与异常值标准化", input: "多源原始字段", output: "标准化业务字段", condition: "执行统一质量规则", calls: "已有能力，被9项产品调用" });
    setNodes(next); setActiveId("clean"); message.success("已增加已有能力：数据清洗");
  };
  const addDataMapping = () => {
    if (nodes.some(n => n.id === "mapping")) return message.info("数据映射模块已在画布中");
    const next = [...nodes];
    next.splice(2, 0, { id: "mapping", name: "数据映射", type: "common", function: "将来源字段转换为统一业务字段", input: "供应商字段与来源口径", output: "标准业务字段", condition: "字段映射关系已通过校验", calls: "已有能力，被11项产品调用" });
    setNodes(next); setActiveId("mapping"); message.success("已增加已有能力：数据映射");
  };
  const libraryModuleId = (name: string) => ({
    "主体关联":"relation","数据映射":"mapping","数据清洗":"clean","半导体链属定位":"industry",
    "扩产事件识别":"event","融资需求研判":"need","风险核验":"risk"
  } as Record<string,string>)[name];
  const addLibraryModule = (id: string) => {
    if (id === "clean") return addModule();
    if (id === "mapping") return addDataMapping();
    const existing = nodes.find(n => n.id === id);
    if (existing) { setActiveId(id); return message.info(`${existing.name}已在画布中，可直接拖动调整顺序`); }
  };
  const reorderCanvasNode = (dragId: string, targetId: string) => {
    if (!dragId || dragId === targetId) return;
    const stages = [["relation","mapping","clean"],["industry"],["event","need"],["risk"],["decision"]];
    const sourceStage = stages.findIndex(s => s.includes(dragId));
    const targetStage = stages.findIndex(s => s.includes(targetId));
    if (sourceStage < 0 || sourceStage !== targetStage) {
      return message.warning("该调整不符合业务阶段，请在同一阶段内调整先后顺序");
    }
    const next = [...nodes];
    const from = next.findIndex(n => n.id === dragId);
    const to = next.findIndex(n => n.id === targetId);
    if (from < 0 || to < 0) return;
    const [picked] = next.splice(from, 1);
    const targetIndex = next.findIndex(n => n.id === targetId);
    next.splice(targetIndex, 0, picked);
    setNodes(next); setActiveId(dragId);
    message.success(`已将“${picked.name}”调整到“${nodes[to].name}”之前`);
  };
  const dropLibraryModule = (e: React.DragEvent, stageIndex: number) => {
    e.preventDefault();
    setDropStage(null);
    const libraryId = e.dataTransfer.getData("library-module");
    if (!libraryId) return;
    const stageById: Record<string,number> = { relation:0, mapping:0, clean:0, industry:1, event:2, need:2, risk:3, decision:4 };
    if (stageById[libraryId] !== stageIndex) return message.warning("该模块与当前阶段不匹配，请拖入对应业务阶段");
    addLibraryModule(libraryId);
  };
  const applyAdapt = () => {
    if (adaptType === "industry") {
      setNodes(nodes.map(n => n.id === "industry" ? { ...n, name: "新能源链属定位", input: "经营范围、产品、资质、项目与新能源供应链信息", output: "电池、材料、整车、储能、设备环节及证据", condition: "新能源产业分类口径" } : n));
      setRulesState({ ...rulesState, industry: "新能源" });
    } else {
      setRulesState({ ...rulesState, region: "安徽省" });
    }
    setAdapted(true); setAdaptOpen(false); message.success("适配版本已生成：共性模块保持复用，差异模块与规则已更新");
  };
  const funnel = running ? [12840, 936, 214, 37, 18] : [12840, 0, 0, 0, 0];
  const branchResources = [
    ["工商登记","EXT-R001","企业名称、注册资本、成立日期"],
    ["企业资质","EXT-R004","专精特新、高新技术、产业资质"],
    ["项目建设","EXT-R007","项目备案、投资额、建设进度"],
    ["招投标","EXT-R003","中标企业、项目金额、发布日期"],
    ["融资事件","EXT-R009","融资轮次、金额、投资机构"],
    ["司法涉诉","EXT-R005","案号、涉案金额、风险等级"],
    ["经营异常","EXT-R006","异常类型、列入及移出日期"],
  ];
  const toggleBranchResource = (name:string) => setBranchSelected(xs => xs.includes(name) ? xs.filter(x=>x!==name) : [...xs,name]);
  return <div className="report-demo-shell">
    <header className="report-demo-header"><BankBrand compact /><div className="report-demo-title"><b>课题汇报演示模式</b><span>半导体企业扩产融资机会与风险分析产品建设</span></div><Space><DemoTag /><Button onClick={onExit}>退出演示</Button></Space></header>
    <div className="report-demo-steps">{["能力模块建设","能力模块组装","标准产品建设","产品应用展示"].map((x,i)=><button key={x} className={step===i?"active":step>i?"done":""} onClick={()=>setStep(i)}><span>{step>i?"✓":i+1}</span><b>{x}</b><small>{["定义、实现与验证","调用、编排与适配","定版、测试与发布","分析、研判与行动"][i]}</small></button>)}</div>
    <main className="report-demo-main">
      {step===0 && <div className="demo-report-page">
        <div className="demo-page-head"><div><Tag color="red">能力模块建设工作台</Tag><h1>扩产融资机会识别能力建设</h1><p>将项目备案、设备招标等经营事实转化为可核实的融资机会与业务证据</p></div><Button size="large" type="primary" onClick={()=>setStep(1)}>完成模块建设，进入组装</Button></div>
        <div className="branch-build-header"><div><small>业务任务</small><b>识别具有潜在融资需求且风险可控的半导体产业客户</b></div><div><small>使用场景</small><b>客户营销</b></div><div><small>适用范围</small><b>江苏地区</b></div><Tag color="orange">地方产品</Tag></div>
        <div className="module-studio">
          <aside className="module-catalog"><div className="catalog-tabs"><button className={catalogType==="resource"?"active":""} onClick={()=>setCatalogType("resource")}>外数资源</button><button className={catalogType==="asset"?"active":""} onClick={()=>setCatalogType("asset")}>模块库</button></div>{catalogType==="resource"?<><div className="catalog-breadcrumb">企业类 › 经营行为与产业信息</div>{[["企业基本信息","工商登记 · 股权关系"],["企业经营行为","项目备案 · 环评 · 设备招标"],["企业融资信息","融资事件 · 授信关系"],["企业风险信息","司法涉诉 · 经营异常"]].map((x,i)=><div className="catalog-level" key={x[0]}><span>0{i+1}</span><div><b>{x[0]}</b><small>{x[1]}</small></div><PlusOutlined /></div>)}</>:<><div className="catalog-breadcrumb">本产品能力模块 · 3项建设中</div>{[["半导体链属定位","输出产业链环节与判断证据"],["扩产事件识别","输出扩产阶段、投资额与证据链"],["融资需求研判","输出需求类型、窗口与建议产品"],...customModules.map(x=>[x,"新增草稿 · 待配置输入输出"])].map((x,i)=><div className={`catalog-level ${i===1?"selected":""}`} key={x[0]}><span>0{i+1}</span><div><b>{x[0]}</b><small>{x[1]}</small></div><CheckCircleFilled /></div>)}<Button block type="dashed" icon={<PlusOutlined />} onClick={()=>setNewModuleOpen(true)}>新建能力模块</Button></>}</aside>
          <section className="module-builder"><header><div><Tag color="red">本次重点建设</Tag><h2>扩产事件识别</h2><p>识别晶圆厂扩产、设备采购和厂房建设等事实，为融资需求研判提供证据</p></div><Space><Button onClick={()=>message.success("模块草稿已保存")}>保存草稿</Button><Button type="primary" onClick={()=>{setStep(1);message.success("能力模块已保存并进入组装环节")}}>完成模块建设，进入组装</Button></Space></header><div className="module-build-tabs">{[["define","1 定义能力"],["contract","2 输入输出"],["implement","3 实现与规则"],["test","4 测试验证"]].map(x=><button key={x[0]} className={modulePanel===x[0]?"active":""} onClick={()=>setModulePanel(x[0] as any)}>{x[1]}</button>)}</div>
            {modulePanel==="define"&&<div className="module-form-grid"><label>能力名称<Input value="扩产事件识别" readOnly/></label><label>能力类型<Select value="经营事件识别" options={[{value:"经营事件识别"}]}/></label><label className="wide">业务语义<Input.TextArea value="识别企业是否正在新建或扩建产线，并说明投资规模、建设阶段及判断依据。" readOnly rows={2}/></label><label>来源产品<Input value="江苏分行半导体客户机会识别" readOnly/></label><label>责任部门<Input value="江苏分行产品创新团队" readOnly/></label><label>适用场景<Select mode="multiple" value={["客户营销","授信支持"]} options={["客户营销","授信支持"].map(value=>({value}))}/></label><label>适用边界<Input value="只识别可验证事件，不直接形成授信结论" readOnly/></label><label className="wide">典型业务信号<div className="module-classification">{["项目备案","环评批复","设备招标","厂房建设","融资事件"].map(x=><span key={x}>{x}<CheckCircleFilled /></span>)}</div></label></div>}
            {modulePanel==="contract"&&<div className="contract-board"><div><h3>标准输入</h3>{[["enterprise_id","企业主体ID","必填"],["project_record","项目备案信息","必填"],["equipment_bid","设备招标信息","必填"],["construction","厂房建设进度","选填"],["finance_event","融资事件","选填"]].map(x=><span key={x[0]}><code>{x[0]}</code><b>{x[1]}</b><Tag color={x[2]==="必填"?"red":"blue"}>{x[2]}</Tag></span>)}</div><em>事实识别与交叉验证 →</em><div><h3>标准输出</h3>{[["expansion_flag","是否存在扩产"],["project_stage","备案／建设／投产阶段"],["investment_amt","投资规模区间"],["evidence_chain","判断依据与来源"],["review_flag","是否需要人工复核"]].map(x=><span key={x[0]}><code>{x[0]}</code><b>{x[1]}</b><CheckCircleFilled /></span>)}</div></div>}
            {modulePanel==="implement"&&<div className="implementation-board"><div className="implement-route">{[["01","事件抽取","从备案、招标和环评文本抽取项目事件"],["02","主体归并","将项目公司、建设单位关联至标准主体"],["03","阶段判断","按备案、开工、设备进场识别建设阶段"],["04","交叉验证","至少两类独立信号相互印证"],["05","证据封装","输出扩产结论、规模区间和来源"]].map((x,i)=><React.Fragment key={x[0]}><span><small>{x[0]}</small><b>{x[1]}</b><em>{x[2]}</em></span>{i<4&&<i>→</i>}</React.Fragment>)}</div><div className="implement-settings"><label>实现方式<Select value="事件规则＋文本抽取服务" options={[{value:"事件规则＋文本抽取服务"}]}/></label><label>调用接口<Input value="POST /capability/event/capex" readOnly/></label><label>冲突处理<Select value="证据不足转人工复核" options={[{value:"证据不足转人工复核"}]}/></label><label>版本策略<Input value="地方试用 V0.9 → 总行评审 → 标准能力 V1.0" readOnly/></label></div></div>}
            {modulePanel==="test"&&<div className="module-test-board"><div className="sample-input"><h3>脱敏样本输入</h3><Descriptions size="small" bordered column={1} items={[{label:"项目备案",children:"12英寸晶圆制造二期项目，计划投资28亿元"},{label:"设备招标",children:"新增光刻、刻蚀及检测设备采购批次"},{label:"建设进度",children:"厂房主体施工，预计次年设备进场"},{label:"融资事件",children:"暂无公开项目融资信息"}]}/></div><div className="module-test-result"><h3>模块输出</h3><div className="test-main-result"><Tag color="red">扩产建设期</Tag><b>证据完整度 92%</b></div><p>结论：存在明确扩产活动。项目备案、设备招标与厂房建设三类信号一致，可进入融资需求研判模块。</p><div className="test-checks"><span><CheckCircleFilled />项目主体一致</span><span><CheckCircleFilled />建设阶段明确</span><span><CheckCircleFilled />投资规模可追溯</span><span><CheckCircleFilled />证据不足转人工</span></div><Button type="primary" onClick={()=>{setModuleTested(true);message.success("12组脱敏样本测试通过，测试记录已写入能力档案")}}>{moduleTested?"测试已通过":"运行12组样本测试"}</Button></div></div>}
          </section>
          <aside className="module-governance"><h3>能力建设要素</h3><p>每项能力必须具备可识别、可组合、可调用和可治理的完整契约。</p>{[["业务语义","解决什么问题"],["数据依赖","调用哪些资源与字段"],["输入输出","上下游如何连接"],["实现逻辑","SQL／Python／API／工作流"],["规则参数","哪些内容可以配置"],["适用边界","何时适用、何时转人工"],["测试记录","用什么样本验证"],["版本责任","谁维护、如何升级"]].map((x,i)=><div key={x[0]}><span>{i+1}</span><b>{x[0]}</b><small>{x[1]}</small><CheckCircleFilled /></div>)}<footer><b>{moduleTested?"8/8 已完成":"7/8 已完成"}</b><Progress percent={moduleTested?100:88} showInfo={false} strokeColor="#c7000b"/><small>{moduleTested?"可提交总行评审":"完成测试后即可提交"}</small></footer></aside>
        </div>
      </div>}
      {false && step===1 && buildMode==="analysis" && <div className="demo-analysis-page">
        <div className="demo-page-head"><div><Tag color="green">结构化比对完成</Tag><h1>标准化分析结果</h1><p>比较业务任务、数据输入、处理节点、规则口径、输出结果和适用范围</p></div><Button size="large" type="primary" onClick={()=>setBuildMode("canvas")}>进入建设画布</Button></div>
        <div className="demo-analysis-columns">
          <section><header><FileSearchOutlined /><b>分行原始产品</b></header><div className="mini-flow">{["多源数据","主体关联","产业识别","特征计算","名单筛选","风险核验","名单输出"].map((x,i)=><React.Fragment key={x}><span>{x}</span>{i<6&&<em>→</em>}</React.Fragment>)}</div><p>完整建设过程已转为统一产品描述，可复现、可比对、可追溯。</p></section>
          <section><header><ApartmentOutlined /><b>共性与差异识别</b></header><h4>已有共性能力</h4><div className="demo-tags common">{["主体关联","数据映射","数据清洗","特征计算","名单筛选","风险核验"].map(x=><span key={x}>{x}</span>)}</div><h4>差异内容</h4><div className="demo-tags diff">{["半导体产业识别","江苏园区名单","地区参数","行业筛选规则"].map(x=><span key={x}>{x}</span>)}</div></section>
          <section><header><DeploymentUnitOutlined /><b>推荐建设路径</b></header><div className="path-statement">复用已有共性能力，新增半导体产业识别模块，配置江苏地区规则，形成标准化产品。</div><div className="path-counts"><span><b>6</b>复用能力</span><span><b>1</b>新增模块</span><span><b>4</b>调整项</span></div></section>
        </div>
        <Card title="相似产品查重 · 相似来自处理环节，不只比较名称"><Table size="small" pagination={false} rowKey="name" dataSource={[{name:"先进制造客户识别",same:"主体关联、特征计算、名单筛选、风险核验",diff:"产业分类与地方规则"},{name:"新能源产业客户筛选",same:"主体关联、特征计算、名单筛选、风险核验",diff:"行业识别模块"}]} columns={[{title:"已有产品",dataIndex:"name"},{title:"相似处理环节",dataIndex:"same",render:(x:string)=><span className="green">{x}</span>},{title:"主要差异",dataIndex:"diff"}]} /></Card>
      </div>}
      {step===1 && <div className="demo-canvas-page">
        <div className="demo-page-head"><div><Tag color="red">标准产品草案 V1.0</Tag><h1>业务链路驱动的模块组装画布</h1><p>模块必须满足输入输出契约和前后依赖；可拖动，但不允许形成无业务含义的流程</p></div><Space><Button icon={<PlusOutlined />} onClick={()=>setNewModuleOpen(true)}>新建模块</Button><Button icon={<SwapOutlined />} onClick={()=>setAdaptOpen(true)}>创建适配版本</Button><Button type="primary" icon={<ExperimentOutlined />} onClick={()=>{setRunning(true);message.success("全链路脱敏样本运行完成")}}>运行预览</Button></Space></div>
        <div className="demo-canvas-layout"><aside><b>模块与资源库</b><small>所有模块均可拖入画布；画布内也可调整顺序</small><h4>数据准备</h4>{["主体关联","数据映射","数据清洗"].map(x=>{const id=libraryModuleId(x);return <button key={x} draggable className="library-draggable" onDragStart={e=>{e.dataTransfer.effectAllowed="copy";e.dataTransfer.setData("library-module",id)}} onClick={()=>addLibraryModule(id)}><DeploymentUnitOutlined /><span>{x}<small>拖入“数据准备”阶段</small></span><PlusOutlined /></button>})}<h4>分析研判</h4>{["半导体链属定位","扩产事件识别","融资需求研判","风险核验"].map(x=>{const id=libraryModuleId(x);return <button key={x} draggable className="library-draggable" onDragStart={e=>{e.dataTransfer.effectAllowed="copy";e.dataTransfer.setData("library-module",id)}} onClick={()=>addLibraryModule(id)}><DeploymentUnitOutlined /><span>{x}<small>拖入对应业务阶段</small></span><PlusOutlined /></button>})}<h4>外数资源</h4>{["项目备案与环评","设备招标与建设进度","司法与经营异常"].map(x=><button key={x} onClick={()=>message.success(`${x}已挂接到兼容节点`)}><DatabaseOutlined /><span>{x}<small>授权与质量校验通过</small></span><PlusOutlined /></button>)}</aside>
          <section className="demo-canvas-center"><div className="hq-business-demand"><small>标准产品业务任务</small><b>发现半导体企业扩产机会，研判融资需求，完成风险校验并形成业务行动</b><span>画布中的每个输出均对应最终驾驶舱的一项分析结果</span></div><div className="demo-change-summary">{adapted?<><Tag color="green">共性模块继续复用</Tag><Tag color="red">行业模块已替换</Tag><Tag color="blue">地区规则已调整</Tag></>:<><Tag color="green">已有能力</Tag><Tag color="red">本次建设</Tag><Tag color="blue">结果映射</Tag><span>拖入模块并在同一阶段内调整先后顺序</span></>}</div><div className="business-pipeline">{[["01 数据准备",nodes.filter(n=>["relation","mapping","clean"].includes(n.id)).map(n=>n.id),"标准主体与标准字段"],["02 产业定位",["industry"],"产业链结构"],["03 机会研判",nodes.filter(n=>["event","need"].includes(n.id)).map(n=>n.id),"经营变化／融资机会"],["04 风险校验",["risk"],"风险关注主体"],["05 决策输出",["decision"],"行动任务"]].map((stage,si)=><React.Fragment key={stage[0] as string}><div className={`pipeline-stage s${si+1} ${dropStage===si?"drop-ready":""}`} onDragEnter={e=>{if(e.dataTransfer.types.includes("library-module"))setDropStage(si)}} onDragOver={e=>{if(e.dataTransfer.types.includes("library-module")){e.preventDefault();e.dataTransfer.dropEffect="copy"}}} onDragLeave={e=>{if(!e.currentTarget.contains(e.relatedTarget as Node))setDropStage(null)}} onDrop={e=>dropLibraryModule(e,si)}><header><b>{stage[0]}</b><small>{dropStage===si?"松开以加入对应阶段":<>对应结果：{stage[2]}</>}</small></header>{(stage[1] as string[]).map(id=>{const n=nodes.find(x=>x.id===id)!;const i=nodes.findIndex(x=>x.id===id);return <button key={id} draggable onDragStart={e=>{e.stopPropagation();e.dataTransfer.effectAllowed="move";e.dataTransfer.setData("canvas-node",id);e.dataTransfer.setData("node",String(i))}} onDragOver={e=>{if(e.dataTransfer.types.includes("canvas-node"))e.preventDefault()}} onDrop={e=>{const dragId=e.dataTransfer.getData("canvas-node");if(dragId){e.preventDefault();e.stopPropagation();reorderCanvasNode(dragId,id)}}} onClick={()=>setActiveId(id)} className={`${n.type} ${activeId===id?"active":""}`} title="拖到同阶段另一模块上方可调整顺序"><DragOutlined /><span><small>{n.type==="common"?"已有能力":"本次建设"}</small><b>{n.name}</b><em>{n.output}</em></span></button>})}</div>{si<4&&<i className="pipeline-arrow">→</i>}</React.Fragment>)}</div><div className="canvas-result-map"><b>组装结果与应用界面一一对应</b>{[["半导体链属定位","产业链结构与机会分布"],["扩产事件识别","经营事件变化趋势"],["融资需求研判","潜在融资机会／预计需求"],["风险核验","区域机会—风险矩阵"],["机会分层与行动建议","业务任务中心"]].map(x=><span key={x[0]}><em>{x[0]}</em><i>→</i>{x[1]}</span>)}</div><div className="demo-funnel-inline">{["全量企业","产业定位","机会识别","风险校验","重点行动"].map((x,i)=><React.Fragment key={x}><span><small>{x}</small><b>{funnel[i].toLocaleString()}</b></span>{i<4&&<em>→</em>}</React.Fragment>)}<Tag>演示数据</Tag></div></section>
          <aside className="demo-inspector"><b>节点属性</b><Tag color={active.type==="new"?"red":active.type==="common"?"green":"blue"}>{active.type==="new"?"本次新增":active.type==="common"?"可复用能力":"流程节点"}</Tag><h3>{active.name}</h3><dl><dt>模块功能</dt><dd>{active.function}</dd><dt>输入数据</dt><dd>{active.input}</dd><dt>输出结果</dt><dd>{active.output}</dd><dt>适用条件</dt><dd>{active.condition}</dd><dt>调用情况</dt><dd>{active.calls}</dd></dl>{active.id==="industry"&&<div className="chain-tags">{(adapted?["电池","材料","整车","储能","设备"]:["设计","制造","封测","设备","材料"]).map(x=><span key={x}>{x}</span>)}</div>}{active.id==="risk"&&<div className="module-build-path"><b>可行建设路径</b><span>接入司法与经营异常资源</span><em>→</em><span>字段映射与事件标准化</span><em>→</em><span>规则封装与脱敏样本验证</span></div>}</aside></div>
        <div className="demo-rule-panel"><div><b>通用规则配置</b><small>修改后预计筛选结果随之变化 · 演示数据</small></div><Select value={rulesState.region} onChange={region=>setRulesState({...rulesState,region})} options={["江苏省","安徽省","浙江省"].map(value=>({value}))}/><Select value={rulesState.industry} onChange={industry=>setRulesState({...rulesState,industry})} options={["半导体","新能源","生物医药"].map(value=>({value}))}/><span>注册年限<InputNumber value={rulesState.years} onChange={years=>setRulesState({...rulesState,years:Number(years)})} /></span><span>经营变化<InputNumber value={rulesState.operation} onChange={operation=>setRulesState({...rulesState,operation:Number(operation)})} addonAfter="月" /></span><span>风险阈值<InputNumber value={rulesState.risk} onChange={risk=>setRulesState({...rulesState,risk:Number(risk)})} /></span><Button onClick={()=>message.info("资质、司法风险与经营异常条件已展开")}>更多条件</Button></div>
        {running&&<div className="demo-run-actions"><span><CheckCircleFilled /> 组装检查通过：资源、模块、规则和输入输出契约均已连接。</span><Button type="primary" onClick={()=>setStep(2)}>固化组装方案，建设产品</Button></div>}
      </div>}
      {step===2 && <div className="product-construction-page"><div className="demo-page-head"><div><Tag color="red">标准产品草案 · BUILD-SEM-01</Tag><h1>半导体企业扩产融资机会与风险分析产品</h1><p>固化“产业定位—扩产识别—需求研判—风险校验—行动建议”完整业务链路</p></div><Button type="primary" size="large" onClick={()=>{setStep(3);message.success("标准产品 V1.0 已发布并进入应用门户")}}>发布产品并查看应用界面</Button></div><div className="product-build-overview"><section><small>业务任务</small><b>发现扩产机会、研判融资需求、核验风险并形成业务行动</b></section><section><small>产品场景</small><b>产业研究 × 客户营销 × 授信支持 × 风险监测</b></section><section><small>适用范围</small><b>江苏试点，可创建地区及行业适配版本</b></section><section><small>发布版本</small><b>V1.0 · 标准产品</b></section></div><div className="product-build-grid"><Card title="产品结构清单"><div className="product-package">{[["外数资源","7项","工商、资质、项目备案、环评、设备招标、司法、经营异常"],["能力模块","6项","主体关联、链属定位、扩产识别、需求研判、风险核验、行动建议"],["通用规则","9项","地区、行业、观察期、证据条件、风险阈值及排除条件"],["决策输出","5类","产业结构、经营变化、融资机会、风险提示、业务任务"]].map((x,i)=><div key={x[0]}><span>0{i+1}</span><div><b>{x[0]}<em>{x[1]}</em></b><small>{x[2]}</small></div><CheckCircleFilled /></div>)}</div></Card><Card title="产品参数与交付配置"><Descriptions size="small" bordered column={1} items={[{label:"默认地区",children:"江苏省（可配置）"},{label:"默认行业",children:"半导体（可替换链属定位模块）"},{label:"证据标准",children:"至少两类独立经营信号交叉验证"},{label:"运行周期",children:"每日增量运行＋月度全量校准"},{label:"交付界面",children:"产业驾驶舱、企业详情、任务中心"},{label:"责任部门",children:"总行数据管理部门＋业务归口部门"}]}/></Card><Card title="验证与发布门禁"><div className="product-release-gates">{[["契约检查","各模块输入输出可连接","通过"],["依赖检查","业务步骤顺序完整","通过"],["样本测试","扩产事实与证据可追溯","通过"],["业务验证","融资线索可由客户经理核实","通过"],["发布审批","地方试点范围","待确认"]].map(x=><div key={x[0]}><CheckCircleFilled /><span><b>{x[0]}</b><small>{x[1]}</small></span><Tag color={x[2]==="通过"?"green":"orange"}>{x[2]}</Tag></div>)}</div></Card></div><div className="product-build-flow">{["业务任务","模块契约","逻辑组装","全链路验证","审批发布","运行评价"].map((x,i)=><React.Fragment key={x}><span className={i<4?"done":i===4?"active":""}><b>{i+1}</b>{x}</span>{i<5&&<em>→</em>}</React.Fragment>)}</div></div>}
      {step===3 && <div className="demo-result-page decision-dashboard"><div className="demo-page-head"><div><Tag color="green">全链路产业经营分析 · 演示数据</Tag><h1>{adapted&&adaptType==="industry"?"新能源":"半导体"}产业客户机会与风险决策驾驶舱</h1><p>{rulesState.region}　｜　产业研究 × 客户营销 × 授信支持 × 风险监测　｜　2026-08-16 14:30</p></div><Space><Button>生成行业简报</Button><Button type="primary" icon={<MessageOutlined />} onClick={()=>setReviewOpen(true)}>应用评价</Button></Space></div><div className="decision-kpis">{[["产业企业","936","家","已完成主体与链属识别"],["重点经营事件","286","项","项目、订单、扩产与融资"],["潜在融资机会","42","项","具备进一步核实价值"],["风险关注主体","17","家","需先核验再开展业务"],["预计融资需求","31.6","亿元","基于事件的演示测算"]].map(x=><div key={x[0]}><small>{x[0]}</small><b>{x[1]}<em>{x[2]}</em></b><span>{x[3]}</span></div>)}</div><div className="decision-conclusion"><RobotOutlined /><div><small>综合研判</small><b>江苏半导体产业投资与订单活动保持活跃，机会主要集中于晶圆制造扩产、设备国产替代和先进封装；部分企业同时出现司法或经营异常信号，应实行“机会识别—风险核验—产品匹配—持续跟踪”的分层推进。</b></div><Tag color="red">人工复核后使用</Tag></div><div className="decision-chart-grid"><Card title="产业链结构与机会分布"><ReactECharts style={{height:245}} option={{tooltip:{trigger:"axis"},legend:{bottom:0},grid:{left:45,right:20,top:25,bottom:45},xAxis:{type:"category",data:["设计","制造","封测","设备","材料"]},yAxis:{type:"value"},series:[{name:"产业企业",type:"bar",data:[318,126,204,158,130],itemStyle:{color:"#d9a5a8"}},{name:"高机会企业",type:"line",data:[11,9,8,9,5],itemStyle:{color:"#c7000b"},lineStyle:{width:3}}]}}/></Card><Card title="区域机会—风险矩阵"><ReactECharts style={{height:245}} option={{tooltip:{},grid:{left:48,right:20,top:25,bottom:40},xAxis:{name:"机会指数",min:40,max:100},yAxis:{name:"风险指数",min:0,max:85},series:[{type:"scatter",symbolSize:(v:any)=>v[2],data:[[88,24,34,"苏州"],[81,32,29,"无锡"],[75,18,25,"南京"],[69,54,22,"常州"],[63,67,18,"南通"]],label:{show:true,formatter:(p:any)=>p.value[3],position:"top"},itemStyle:{color:(p:any)=>p.value[1]>50?"#d77c00":"#c7000b"},markLine:{silent:true,data:[{xAxis:70},{yAxis:50}]}}]}}/></Card><Card title="经营事件变化趋势"><ReactECharts style={{height:245}} option={{tooltip:{trigger:"axis"},legend:{bottom:0},grid:{left:45,right:18,top:25,bottom:45},xAxis:{type:"category",data:["3月","4月","5月","6月","7月","8月"]},yAxis:{type:"value"},series:[{name:"项目/扩产",type:"line",smooth:true,areaStyle:{opacity:.12},data:[18,22,31,29,38,46],itemStyle:{color:"#c7000b"}},{name:"订单/招标",type:"line",smooth:true,data:[26,31,28,39,44,51],itemStyle:{color:"#315d88"}},{name:"风险事件",type:"bar",data:[8,7,11,9,13,10],itemStyle:{color:"#d77c00"}}]}}/></Card></div><div className="decision-lower-grid"><Card title="从事实到业务行动的证据链"><div className="evidence-decision-chain">{[["事实信号","新增产线备案、设备招标、融资事件"],["产业判断","晶圆制造与设备环节资本开支上升"],["需求推断","项目建设、设备采购与周转融资需求"],["风险反证","司法涉诉、经营异常与行业风险核验"],["业务行动","优先访谈、方案匹配、授信核验、持续监测"]].map((x,i)=><React.Fragment key={x[0]}><div><span>0{i+1}</span><b>{x[0]}</b><small>{x[1]}</small></div>{i<4&&<em>→</em>}</React.Fragment>)}</div></Card><Card title="产业链重点机会"><div className="chain-opportunities">{[["晶圆制造扩产","9家","项目贷款、设备融资"],["设备国产替代","9家","订单融资、供应链金融"],["先进封装升级","8家","技改贷款、流动资金"],["关键材料扩能","5家","项目融资、结算服务"]].map(x=><div key={x[0]}><span/><b>{x[0]}</b><em>{x[1]}</em><small>{x[2]}</small></div>)}</div></Card></div><Card title="业务任务中心 · 名单只是行动明细" className="decision-task-center"><div>{[["客户营销","18","优先访谈","按属地推送客户经理"],["融资方案","12","待匹配","项目贷、设备贷、供应链融资"],["授信核验","7","待协同","核验项目真实性与资金用途"],["风险监测","17","持续观察","司法、异常和关联风险跟踪"]].map(x=><button key={x[0]} onClick={x[0]==="客户营销"?()=>setCompanyOpen(true):()=>message.success(`${x[0]}任务已进入待办中心`)}><span className="task-number">{x[1]}</span><div><b>{x[0]}</b><small>{x[3]}</small></div><Tag className="task-status" color={x[2]==="持续观察"?"orange":"red"}>{x[2]}</Tag></button>)}</div><footer><span>筛选漏斗：12,840家全量企业 → 936家产业企业 → 214家条件入围 → 37家风险筛除 → 18家重点跟进</span><Button size="small" onClick={()=>setCompanyOpen(true)}>查看客户明细</Button></footer></Card><div className="demo-loop"><span>产品发布</span><em>→</em><span>全链路分析</span><em>→</em><span>业务协同</span><em>→</em><span>评价反馈</span><em>→</em><span>定位产品／模块／规则</span><em>→</em><span>版本优化</span></div></div>}
    </main>
    <Modal open={newModuleOpen} onCancel={()=>setNewModuleOpen(false)} title="新建能力模块" okText="创建模块草稿" onOk={()=>{if(!newModuleName.trim())return message.warning("请填写模块名称");setCustomModules([...customModules,newModuleName.trim()]);setNewModuleName("");setNewModuleOpen(false);setCatalogType("asset");message.success("模块草稿已创建，请继续配置业务语义和输入输出")}}><Form layout="vertical"><Form.Item label="模块名称" required><Input value={newModuleName} onChange={e=>setNewModuleName(e.target.value)} placeholder="例如：设备采购资金需求研判" /></Form.Item><Form.Item label="所属阶段"><Select defaultValue="机会研判" options={["数据准备","产业定位","机会研判","风险校验","决策输出"].map(value=>({value}))}/></Form.Item><Form.Item label="建设要求"><div className="new-module-hint">创建后需依次配置业务语义、标准输入输出、实现规则、适用边界和测试记录；只有契约校验通过后才能加入画布。</div></Form.Item></Form></Modal>
    <Modal open={adaptOpen} onCancel={()=>setAdaptOpen(false)} title="创建适配版本" okText="生成适配版本" onOk={applyAdapt}><Radio.Group value={adaptType} onChange={e=>setAdaptType(e.target.value)}><Space direction="vertical"><Radio value="region"><b>地区适配</b>：江苏替换为安徽，同步替换园区名单与地区规则</Radio><Radio value="industry"><b>行业适配</b>：半导体链属定位替换为新能源链属定位，其余业务模块继续复用</Radio></Space></Radio.Group><div className="adapt-impact"><span><b>5</b>共性模块继续复用</span><span><b>1</b>行业模块被替换</span><span><b>1</b>规则组调整</span><span><b>1</b>地方名单需新增</span></div></Modal>
    <Drawer width={520} open={companyOpen} onClose={()=>setCompanyOpen(false)} title="苏芯微电子有限公司 · 单户详情"><div className="company-explain"><Tag color="red">重点跟进 · 演示数据</Tag><h2>为什么被识别为目标客户？</h2><ol><li><b>产业位置明确：</b>经营范围、产品信息和资质信息共同指向芯片设计环节。</li><li><b>经营变化可验证：</b>近期新增研发中心建设项目，并连续发布设备采购招标。</li><li><b>潜在需求合理：</b>项目建设和设备采购可能形成中长期项目融资及流动资金需求。</li><li><b>风险处于可核验范围：</b>发现1项一般合同纠纷，未达到自动排除阈值，建议客户经理访谈前核验。</li></ol><Descriptions bordered size="small" column={1} items={[{label:"入围依据",children:"新增研发中心＋设备采购项目＋有效产业资质"},{label:"经营变化",children:"近6个月新增招标3项，项目建设进度更新"},{label:"可能需求",children:"项目建设融资、设备融资、流动资金"},{label:"建议动作",children:"由属地客户经理优先访谈，并同步核验涉诉事项"}]} /></div></Drawer>
    <Modal open={reviewOpen} onCancel={()=>setReviewOpen(false)} title="应用评价与优化反馈" okText="提交评价" onOk={()=>{setReviewOpen(false);message.success("评价已关联至产品、模块和规则，返回建设管理端")}}><Form layout="vertical"><Form.Item label="结果是否准确"><Rate defaultValue={4}/></Form.Item><Form.Item label="反馈内容"><Checkbox.Group options={["具有营销价值","存在遗漏","规则需要调整","适合向其他地区推广"]} defaultValue={["具有营销价值","适合向其他地区推广"]}/></Form.Item><Form.Item label="补充说明"><Input.TextArea defaultValue="建议继续观察设备采购事件，并将园区名单按季度更新。" rows={3}/></Form.Item></Form></Modal>
  </div>;
}

function TopHeader({
  role,
  onRole,
  onLogout,
  portal,
  onPortal,
}: {
  role: RoleKey;
  onRole: (r: RoleKey) => void;
  onLogout: () => void;
  portal: Portal;
  onPortal: (p: Portal) => void;
}) {
  return (
    <Header className="top-header">
      <BankBrand compact />
      <Input
        className="global-search"
        prefix={<SearchOutlined />}
        placeholder="搜索产品、资源、能力、规则或业务场景"
        suffix={<span>高级检索</span>}
      />
      <Space size={20} className="header-actions">
        <Badge count={6} size="small">
          <BellOutlined />
        </Badge>
        <StarOutlined />
        <HistoryOutlined />
        <Dropdown
          menu={{
            items: [
              {
                key: "business",
                label: "切换角色：业务人员",
                onClick: () => onRole("business"),
              },
              {
                key: "branch-role",
                label: "切换角色：分行产品人员",
                onClick: () => onRole("branch"),
              },
              {
                key: "admin",
                label: "切换角色：总行管理人员",
                onClick: () => onRole("admin"),
              },
              { type: "divider" },
              {
                key: "app",
                label: "进入业务应用门户",
                onClick: () => onPortal("app"),
              },
              {
                key: "branch",
                label: "进入分行建设与报送门户",
                onClick: () => onPortal("branch"),
              },
              {
                key: "manage",
                label: "进入总行建设管理门户",
                onClick: () => onPortal("manage"),
              },
              {
                key: "logout",
                icon: <LogoutOutlined />,
                label: "退出演示",
                onClick: onLogout,
              },
            ],
          }}
        >
          <div className="user-chip">
            <Avatar size={34} icon={<UserOutlined />} />
            <span>
              <b>{roleMap[role].name}</b>
              <small>{roleMap[role].org}</small>
            </span>
          </div>
        </Dropdown>
      </Space>
    </Header>
  );
}

function MetricCard({
  title,
  value,
  suffix,
  trend,
  alert,
  onClick,
}: {
  title: string;
  value: number | string;
  suffix?: string;
  trend?: string;
  alert?: boolean;
  onClick?: () => void;
}) {
  return (
    <Card className={"metric-card " + (alert ? "alert" : "")} onClick={onClick}>
      <Statistic title={title} value={value} suffix={suffix} />
      {trend && <span className="metric-trend">{trend}</span>}
    </Card>
  );
}

function ManageHome({ go }: { go: (v: View) => void }) {
  const liveDate = useLiveChineseDate();
  const flow = [
    ["需求报送", 8, 1],
    ["完整性校验", 3, 0],
    ["场景映射", 5, 1],
    ["产品查重", 4, 1],
    ["方案设计", 6, 1],
    ["建设开发", 7, 2],
    ["测试验证", 4, 1],
    ["审批发布", 2, 0],
    ["能力回沉", 3, 0],
  ];
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            总行管理首页 <DemoTag />
          </h1>
          <p>全行外部数据产品受理、建设、发布与运行总览 · {liveDate}</p>
        </div>
        <Space>
          <Button>导出月报</Button>
          <Button type="primary" onClick={() => go("candidates")}>
            受理分行报送
          </Button>
        </Space>
      </div>
      <Row gutter={[14, 14]} className="metrics-grid">
        <Col span={4}>
          <MetricCard title="外数资源总数" value={1268} trend="本月新增 18" />
        </Col>
        <Col span={4}>
          <MetricCard title="正式产品" value={86} trend="全行 49 · 地方 31" />
        </Col>
        <Col span={4}>
          <MetricCard
            title="候选产品"
            value={23}
            trend="本月报送 8"
            alert
            onClick={() => go("candidates")}
          />
        </Col>
        <Col span={4}>
          <MetricCard title="能力模块" value={64} trend="复用率 76.8%" />
        </Col>
        <Col span={4}>
          <MetricCard
            title="待验证产品"
            value={4}
            trend="逾期 1"
            alert
            onClick={() => go("build-testing")}
          />
        </Col>
        <Col span={4}>
          <MetricCard
            title="待处理反馈"
            value={12}
            trend="高优先级 3"
            alert
            onClick={() => go("operations")}
          />
        </Col>
      </Row>
      <Row gutter={[14, 14]} className="factory-home-overview">
        <Col span={15}>
          <Card title="标准产品建设主线" extra={<Tag color="red">创新工厂总览</Tag>}>
            <div className="factory-flow">
              <div>
                {["分行报送", "受理定义", "复用设计", "模块组装", "验证发布", "运营回沉"].map((x, i) => (
                  <React.Fragment key={x}>
                    <button onClick={() => go(i === 0 ? "candidates" : i < 4 ? "build" : i === 4 ? "finished" : "operations")}>
                      <span>0{i + 1}</span><b>{x}</b><em>{[8, 5, 6, 7, 4, 12][i]}项</em>
                      <small>{["地方成果进入总行", "形成需求基线", "查重与路径选择", "能力规则流程组合", "测试试点与审批", "评价反馈与资产回沉"][i]}</small>
                    </button>
                    {i < 5 && <i>→</i>}
                  </React.Fragment>
                ))}
              </div>
            </div>
          </Card>
        </Col>
        <Col span={9}>
          <Card title="建设效能">
            <ReactECharts style={{ height: 205 }} option={{
              tooltip: { trigger: "axis" },
              legend: { data: ["建设周期", "复用率"], top: 0, left: "center", itemGap: 24 },
              grid: { left: 42, right: 38, top: 48, bottom: 30, containLabel: true },
              xAxis: { type: "category", data: ["3月", "4月", "5月", "6月", "7月", "8月"] },
              yAxis: [{ type: "value" }, { type: "value", max: 100 }],
              series: [
                { name: "建设周期", type: "bar", data: [28,25,22,19,16,13.6], itemStyle: { color: "#c7000b" } },
                { name: "复用率", type: "line", yAxisIndex: 1, data: [52,57,61,68,73,78], itemStyle: { color: "#d89614" } }
              ]
            }} />
          </Card>
        </Col>
      </Row>
      <Row gutter={14}>
        <Col span={15}>
          <Card>
            <SectionTitle
              title="重点待办"
              extra={<Button type="link">查看全部</Button>}
            />
            <Table
              size="small"
              pagination={false}
              dataSource={[
                {
                  key: 1,
                  item: "受理浙江分行专精特新企业筛选产品",
                  type: "候选产品受理",
                  owner: "邓宇轩",
                  limit: "今天 17:00",
                  status: "待处理",
                },
                {
                  key: 2,
                  item: "确认广东供应链风险产品数据映射",
                  type: "建设评审",
                  owner: "陈嘉敏",
                  limit: "08-01",
                  status: "处理中",
                },
                {
                  key: 3,
                  item: "审批江苏招投标客户拓展工具",
                  type: "发布审批",
                  owner: "王  琳",
                  limit: "08-02",
                  status: "待处理",
                },
                {
                  key: 4,
                  item: "处置舆情风险预警低评分反馈",
                  type: "运营优化",
                  owner: "赵  谦",
                  limit: "08-04",
                  status: "待分派",
                },
              ]}
              columns={[
                {
                  title: "待办事项",
                  dataIndex: "item",
                  render: (t: string) => (
                    <a
                      onClick={() =>
                        go(t.includes("浙江") ? "candidates" : "build")
                      }
                    >
                      {t}
                    </a>
                  ),
                },
                { title: "事项类型", dataIndex: "type" },
                { title: "负责人", dataIndex: "owner" },
                { title: "期限", dataIndex: "limit" },
                {
                  title: "状态",
                  dataIndex: "status",
                  render: (t: string) => <StatusTag status={t} />,
                },
              ]}
            />
          </Card>
        </Col>
        <Col span={9}>
          <Card>
            <SectionTitle title="能力复用概况" />
            <ReactECharts
              style={{ height: 230 }}
              option={{
                tooltip: {},
                grid: { left: 40, right: 20, top: 16, bottom: 30 },
                xAxis: {
                  type: "category",
                  data: [
                    "主体关联",
                    "名单筛选",
                    "司法核验",
                    "企业画像",
                    "资质识别",
                  ],
                },
                yAxis: { type: "value" },
                series: [
                  {
                    type: "bar",
                    data: [128, 112, 88, 61, 57],
                    itemStyle: { color: "#c7000b", borderRadius: [3, 3, 0, 0] },
                  },
                ],
              }}
            />
          </Card>
        </Col>
      </Row>
    </div>
  );
}

function BranchHome({
  go,
  submitted,
  published,
}: {
  go: (v: View) => void;
  submitted: boolean;
  published: boolean;
}) {
  const status = published ? "已上架" : submitted ? "总行建设中" : "草稿完善中";
  return (
    <div className="branch-portal">
      <div className="page-heading">
        <div>
          <h1>
            分行产品工作台 <DemoTag />
          </h1>
          <p>
            面向各分行创新团队，统一开展产品组建、试运行、在线报送和上架跟踪
          </p>
        </div>
        <Button
          type="primary"
          icon={<PlusOutlined />}
          onClick={() => go("branch-build")}
        >
          组建新产品
        </Button>
      </div>
      <div className="branch-linkage">
        {[
          ["01", "分行建设与试运行", "调用统一资源、能力与规则组装产品"],
          ["02", "在线报送总行", "提交运行效果、案例与标准化建议"],
          ["03", "总行标准化建设", "完成查重、优化、全行适配与验证"],
          ["04", "全行应用与反馈", "上架正式产品并回传运营结果"],
        ].map((x, i) => (
          <React.Fragment key={x[0]}>
            <div>
              <span>{x[0]}</span>
              <b>{x[1]}</b>
              <small>{x[2]}</small>
            </div>
            {i < 3 && <em>→</em>}
          </React.Fragment>
        ))}
      </div>
      <Row gutter={[14, 14]} className="metrics-grid branch-metrics">
        <Col xs={24} sm={12} xl={6}>
          <MetricCard
            title="在建产品"
            value={3}
            trend="其中1项已具备报送条件"
          />
        </Col>
        <Col xs={24} sm={12} xl={6}>
          <MetricCard title="累计报送" value={7} trend="2项已进入总行建设" />
        </Col>
        <Col xs={24} sm={12} xl={6}>
          <MetricCard
            title="成功上架"
            value={published ? 3 : 2}
            trend={published ? "本月新增1项" : "另有1项待发布"}
          />
        </Col>
        <Col xs={24} sm={12} xl={6}>
          <MetricCard
            title="近30日反馈"
            value={18}
            trend="已上架产品平均评分4.7分"
          />
        </Col>
      </Row>
      <Row gutter={[14, 14]}>
        <Col xs={24} xl={17}>
          <Card
            title="重点产品进展"
            extra={
              <Button type="link" onClick={() => go("branch-submissions")}>
                查看全部报送
              </Button>
            }
          >
            <Table
              scroll={{ x: 980 }}
              pagination={false}
              rowKey="id"
              dataSource={[
                {
                  id: "BR-SC-2026-017",
                  name: "某分行半导体产业链客户机会识别",
                  stage: status,
                  hq: "全国产业链机会与风险识别",
                  scope: published ? "全行" : "拟标准化",
                  feedback: published
                    ? "总行已重新建设全国标准产品，某分行样本贡献已留痕"
                    : "地方名单产品已具备标准化评估条件",
                },
                {
                  id: "BR-SC-2026-012",
                  name: "成渝重大项目融资线索监测",
                  stage: "总行初评",
                  hq: "区域重大项目融资机会母版",
                  scope: "区域",
                  feedback: "建议与现有重大项目母版合并",
                },
                {
                  id: "BR-SC-2026-009",
                  name: "四川制造业供应链风险监测",
                  stage: "试运行",
                  hq: "—",
                  scope: "地方",
                  feedback: "累计形成风险提示84条",
                },
              ]}
              columns={[
                {
                  title: "分行产品",
                  dataIndex: "name",
                  render: (x: string, r: any) => (
                    <>
                      <b>{x}</b>
                      <small className="cell-sub">{r.id}</small>
                    </>
                  ),
                },
                {
                  title: "当前阶段",
                  dataIndex: "stage",
                  render: (x: string) => <StatusTag status={x} />,
                },
                { title: "总行标准化成果", dataIndex: "hq" },
                { title: "拟适用范围", dataIndex: "scope" },
                { title: "最新反馈", dataIndex: "feedback" },
                {
                  title: "操作",
                  render: (_: any, r: any) => (
                    <Button
                      size="small"
                      onClick={() =>
                        go(
                          r.stage === "试运行"
                            ? "branch-build"
                            : "branch-submissions",
                        )
                      }
                    >
                      查看
                    </Button>
                  ),
                },
              ]}
            />
          </Card>
        </Col>
        <Col xs={24} xl={7}>
          <Card className="branch-todo" title="待办与提醒">
            <Timeline
              items={[
                {
                  color: "red",
                  children: (
                    <>
                      <b>完善运行成效佐证</b>
                      <p>产业事件产品 · 今天17:00前</p>
                    </>
                  ),
                },
                {
                  color: "blue",
                  children: (
                    <>
                      <b>确认总行标准化命名</b>
                      <p>保留来源产品与正式产品映射</p>
                    </>
                  ),
                },
                {
                  color: "green",
                  children: (
                    <>
                      <b>查看正式产品首期反馈</b>
                      <p>全行用户评价已同步至分行端</p>
                    </>
                  ),
                },
              ]}
            />
            <Space direction="vertical" style={{ width: "100%" }}>
              <Button block onClick={() => go("report")}>
                继续在线报送
              </Button>
              <Button block onClick={() => go("branch-feedback")}>
                查看运营反馈
              </Button>
            </Space>
          </Card>
        </Col>
      </Row>
    </div>
  );
}

function BranchProductStudio({ go }: { go: (v: View) => void }) {
  const [selectedResources, setSelectedResources] = useState([
    "EXT-R001",
    "EXT-R003",
    "EXT-R004",
    "EXT-R005",
    "EXT-R008",
    "EXT-R015",
  ]);
  const [selectedAbilities, setSelectedAbilities] = useState([
    "主体关联",
    "产业链节点识别",
    "产业资质识别",
    "招投标活跃度计算",
    "司法风险核验",
    "名单筛选",
  ]);
  const [ran, setRan] = useState(false);
  const [branchStage, setBranchStage] = useState(1);
  const [branchFlow, setBranchFlow] = useState([
    "企业主体输入",
    "主体关联",
    "工商标准化",
    "产业资质核验",
    "招投标活跃度计算",
    "司法风险核验",
    "名单筛选",
    "行动提示",
  ]);
  const [branchSelected, setBranchSelected] = useState(3);
  const resourceOptions = resources
    .slice(0, 12)
    .map((r: any) => ({ label: `${r.name}（${r.id}）`, value: r.id }));
  const abilityOptions = [
    "事件抽取与标准化",
    "主体关联",
    "动态状态更新",
    "多维价值评估",
    "事件影响归因",
    "风险交叉核验",
    "趋势预测",
    "策略推荐与解释",
    "任务路由",
  ];
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            在线组建与试运行 <DemoTag />
          </h1>
          <p>分行可调用总行统一资源和能力形成产品，运行验证后直接发起报送</p>
        </div>
        <Space>
          <Button onClick={() => message.success("方案已保存为分行草稿")}>
            保存草稿
          </Button>
          <Button type="primary" onClick={() => go("report")}>
            转为报送材料
          </Button>
        </Space>
      </div>
      <div className="branch-lifecycle">
        {[
          [1, "场景与任务", "明确对象、流程位置和交付成果"],
          [2, "结构化组建", "调用资源、能力与规则"],
          [3, "试运行验证", "用脱敏样本检查结果"],
          [4, "在线报送", "完整记录建设过程"],
          [5, "总行评审反馈", "查重、标准化与推广"],
        ].map((x, i) => <React.Fragment key={x[0]}><button className={branchStage === Number(x[0]) ? "active" : branchStage > Number(x[0]) ? "done" : ""} onClick={() => setBranchStage(Number(x[0]))}><span>{branchStage > Number(x[0]) ? <CheckCircleFilled /> : x[0]}</span><b>{x[1]}</b><small>{x[2]}</small></button>{i < 4 && <em>→</em>}</React.Fragment>)}
      </div>
      <Card className="branch-flow-card" title="分行产品组装画布" extra={<Space><Tag color="blue">地方试建</Tag><Button size="small" icon={<PlusOutlined />} onClick={() => { setBranchFlow([...branchFlow.slice(0, -1), "地方产业目录映射", branchFlow[branchFlow.length - 1]]); message.success("已增加地方扩展节点"); }}>增加地方节点</Button></Space>}>
        <div className="branch-flow-canvas">
          {branchFlow.map((name, i) => <React.Fragment key={`${name}-${i}`}><button className={branchSelected === i ? "active" : ""} onClick={() => setBranchSelected(i)}><small>{i === 0 ? "输入" : i === branchFlow.length - 1 ? "输出" : `模块 ${String(i).padStart(2, "0")}`}</small><b>{name}</b><span>{name.includes("地方") ? "本地扩展" : i === 0 || i === branchFlow.length - 1 ? "产品节点" : "调用总行能力"}</span></button>{i < branchFlow.length - 1 && <em>→</em>}</React.Fragment>)}
        </div>
        <div className="branch-node-editor"><div><b>当前节点：{branchFlow[branchSelected]}</b><span>分行可配置地方数据映射、适用地区、行业、期限和阈值；总行标准模块内部逻辑保持锁定。</span></div><Select style={{ width: 230 }} value={branchFlow[branchSelected]} onChange={(value) => { const nextFlow = [...branchFlow]; nextFlow[branchSelected] = value; setBranchFlow(nextFlow); }} options={["主体关联","工商标准化","产业资质核验","招投标活跃度计算","司法风险核验","名单筛选","行动提示","地方产业目录映射"].map(value => ({ value }))} /><Button onClick={() => message.success("节点配置已保存并写入报送记录")}>保存节点配置</Button></div>
      </Card>
      <Row gutter={14}>
        <Col span={9}>
          <Card title="产品组装配置">
            <Form layout="vertical">
              <Form.Item label="产品母版">
                <Select
                  defaultValue="地方产业名单筛选母版"
                  options={[
                    "地方产业名单筛选母版",
                    "产业事件机会与风险决策母版",
                    "区域景气监测母版",
                    "客户经营趋势预测母版",
                    "供应链风险监测母版",
                  ].map((value) => ({ value }))}
                />
              </Form.Item>
              <Form.Item label="产品名称">
                <Input defaultValue="某分行半导体产业链客户机会识别" />
              </Form.Item>
              <Form.Item label="分行辖内地区"><Select mode="multiple" defaultValue={["核心城区","重点开发区","产业园区"]} options={["核心城区","重点开发区","产业园区","高新区","经开区"].map(value=>({value}))}/></Form.Item>
              <Form.Item label="半导体产业链环节"><Select mode="multiple" defaultValue={["芯片设计","晶圆制造","封装测试"]} options={["半导体材料","设备零部件","芯片设计","晶圆制造","封装测试","功率器件","第三代半导体"].map(value=>({value}))}/></Form.Item>
              <Row gutter={8}><Col span={12}><Form.Item label="成立年限"><Select defaultValue="近10年" options={["近3年","近5年","近10年","不限"].map(value=>({value}))}/></Form.Item></Col><Col span={12}><Form.Item label="企业规模"><Select defaultValue="中小企业" options={["大型企业","中型企业","中小企业","不限"].map(value=>({value}))}/></Form.Item></Col></Row>
              <Form.Item label="调用外数资源">
                <Select
                  mode="multiple"
                  maxTagCount="responsive"
                  value={selectedResources}
                  onChange={setSelectedResources}
                  options={resourceOptions}
                />
              </Form.Item>
              <Form.Item label="调用能力模块">
                <Select
                  mode="multiple"
                  maxTagCount="responsive"
                  value={selectedAbilities}
                  onChange={setSelectedAbilities}
                  options={abilityOptions.map((value) => ({ value }))}
                />
              </Form.Item>
              <Form.Item label="地方规则">
                <Select
                  mode="multiple"
                  defaultValue={[
                    "分行辖内地域范围",
                    "事件近90日",
                    "重大处罚否决",
                    "机会A级触发营销",
                  ]}
                  options={[
                    "分行辖内地域范围",
                    "事件近90日",
                    "重大处罚否决",
                    "机会A级触发营销",
                    "高风险转人工复核",
                    "存量客户优先路由",
                  ].map((value) => ({ value }))}
                />
              </Form.Item>
              <Button
                type="primary"
                block
                loading={false}
                onClick={() => {
                  setRan(true);
                  setBranchStage(3);
                  message.success("试运行完成，已生成60条演示结果");
                }}
              >
                运行产品
              </Button>
            </Form>
          </Card>
        </Col>
        <Col span={15}>
          <Card
            title="试运行结果"
            extra={
              <Tag color={ran ? "green" : "default"}>
                {ran ? "运行完成" : "等待运行"}
              </Tag>
            }
          >
            {ran ? (
              <>
                <div className="preview-output-grid">
                  <div>
                    <span>识别企业</span>
                    <b>60家</b>
                  </div>
                  <div>
                    <span>融资机会</span>
                    <b>38条</b>
                  </div>
                  <div>
                    <span>风险提示</span>
                    <b>12条</b>
                  </div>
                  <div>
                    <span>待人工核验</span>
                    <b>10条</b>
                  </div>
                </div>
                <Table
                  size="small"
                  pagination={{ pageSize: 5 }}
                  rowKey="name"
                  dataSource={companyResults.slice(0, 12)}
                  columns={[
                    { title: "企业名称", dataIndex: "name" },
                    { title: "地区", dataIndex: "city" },
                    { title: "评分", dataIndex: "score" },
                    {
                      title: "风险等级",
                      dataIndex: "risk",
                      render: (x: string) => (
                        <Tag color={x?.includes("高") ? "red" : "green"}>
                          {x || "低"}
                        </Tag>
                      ),
                    },
                    { title: "业务动作", render: () => "推送客户经理核验" },
                  ]}
                />
                <div className="pane-actions">
                  <Button
                    icon={<DownloadOutlined />}
                    onClick={() =>
                      exportExcel(companyResults, "某分行半导体产品试运行结果.xls")
                    }
                  >
                    导出结果
                  </Button>
                  <Button type="primary" onClick={() => go("report")}>
                    确认效果并报送
                  </Button>
                </div>
              </>
            ) : (
              <Empty description="完成资源、能力和规则配置后运行产品；运行结果及成效指标可自动带入报送表单" />
            )}
          </Card>
        </Col>
      </Row>
    </div>
  );
}

function BranchSubmissions({
  published,
  go,
}: {
  published: boolean;
  go: (v: View) => void;
}) {
  const mainStatus = published ? "已上架" : "总行建设中";
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            我的报送 <DemoTag />
          </h1>
          <p>查看分行报送记录、总行处理进度、标准化改造结果和正式上架情况</p>
        </div>
        <Button type="primary" onClick={() => go("report")}>
          新增报送
        </Button>
      </div>
      <Card>
        <Table
          rowKey="id"
          pagination={false}
          dataSource={[
            {
              id: "CP-2026-0801-09",
              branch: "某分行半导体产业链客户机会识别",
              date: "2026-08-01",
              status: mainStatus,
              hq: "全国产业链机会与风险识别",
              version: published ? "V3.1" : "建设版本 V2.6",
              scope: "全行",
              note: published
                ? "总行已完成优化和全行适配，来源贡献已留痕"
                : "正在补充动态评估、内外数融合和全行参数",
            },
            {
              id: "CP-2026-0721-08",
              branch: "成渝重大项目融资线索监测",
              date: "2026-07-21",
              status: "总行初评",
              hq: "区域重大项目融资机会母版",
              version: "方案设计",
              scope: "区域",
              note: "建议能力组合并保留成渝参数",
            },
            {
              id: "CP-2026-0712-02",
              branch: "四川专精特新营销名单",
              date: "2026-07-12",
              status: "已上架",
              hq: "专精特新企业营销名单",
              version: "V2.4",
              scope: "全行",
              note: "已累计被18家分行复用",
            },
          ]}
          columns={[
            {
              title: "分行报送产品",
              dataIndex: "branch",
              render: (x: string, r: any) => (
                <>
                  <b>{x}</b>
                  <small className="cell-sub">
                    {r.id} · 报送于 {r.date}
                  </small>
                </>
              ),
            },
            {
              title: "总行处理状态",
              dataIndex: "status",
              render: (x: string) => <StatusTag status={x} />,
            },
            { title: "形成的标准产品/母版", dataIndex: "hq" },
            { title: "当前版本", dataIndex: "version" },
            { title: "适用范围", dataIndex: "scope" },
            { title: "总行反馈", dataIndex: "note" },
            {
              title: "操作",
              render: () => (
                <Button size="small" onClick={() => go("branch-feedback")}>
                  查看联动详情
                </Button>
              ),
            },
          ]}
        />
      </Card>
    </div>
  );
}

function BranchFeedback({ published }: { published: boolean }) {
  const feedbackProducts = [
    {
      id: "CP-2026-0801-09",
      source: "某分行半导体产业链客户机会识别",
      org: "某分行",
      formal: "全国产业链机会与风险识别",
      formalId: "P-2026-089",
      version: published ? "V3.1 已发布" : "V3.1 待发布",
      status: published ? "已上架" : "待发布",
      coverage: published ? 18 : 0,
      runs: published ? 1268 : 0,
      rating: published ? 4.7 : 0,
      trend: [186, 278, 356, 448],
      adoption: [62, 68, 73, 77],
      feedback: [
        "保留来源分行的产品创意与首批验证贡献",
        "事件识别和双任务分派能力已纳入全行母版",
        "新增动态机会、风险和客户关系三维评估",
      ],
    },
    {
      id: "CP-2026-0712-02",
      source: "浙江分行专精特新企业营销名单",
      org: "浙江省分行",
      formal: "专精特新企业营销名单",
      formalId: "P-2026-063",
      version: "V2.4 已发布",
      status: "已上架",
      coverage: 24,
      runs: 2864,
      rating: 4.8,
      trend: [540, 686, 762, 876],
      adoption: [68, 72, 75, 81],
      feedback: [
        "地方资质名单已映射为全行标准标签",
        "补充招投标活跃度和司法风险核验",
        "已被24家机构调用并形成营销任务",
      ],
    },
    {
      id: "CP-2026-0721-08",
      source: "广东分行供应链风险监测",
      org: "广东省分行",
      formal: "供应链关联风险监测",
      formalId: "P-2026-071",
      version: "V1.8 灰度运行",
      status: "灰度上架",
      coverage: 8,
      runs: 936,
      rating: 4.5,
      trend: [148, 205, 267, 316],
      adoption: [55, 61, 66, 70],
      feedback: [
        "已完成核心企业与上下游主体映射",
        "风险传导规则进入灰度验证",
        "建议补充跨区域供应链覆盖",
      ],
    },
    {
      id: "CP-2026-0628-04",
      source: "江苏分行招投标客户拓展工具",
      org: "江苏省分行",
      formal: "招投标客户发现与营销推荐",
      formalId: "P-2026-052",
      version: "V2.1 已发布",
      status: "已上架",
      coverage: 20,
      runs: 2146,
      rating: 4.6,
      trend: [392, 508, 574, 672],
      adoption: [60, 65, 69, 74],
      feedback: [
        "统一中标金额、采购人和项目标签口径",
        "新增存量客户排重及管户路由能力",
        "重点反馈为提升项目更新及时性",
      ],
    },
  ];
  const [selectedId, setSelectedId] = useState(feedbackProducts[0].id);
  const item = feedbackProducts.find((x) => x.id === selectedId)!;
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            上架与运营反馈 <DemoTag />
          </h1>
          <p>按报送产品查看正式上架结果、来源贡献、运营成效和总行优化意见</p>
        </div>
      </div>
      <Card className="feedback-selector">
        <div>
          <b>选择报送产品</b>
          <span>共 {feedbackProducts.length} 项已进入上架或运营阶段</span>
        </div>
        <Select
          value={selectedId}
          onChange={setSelectedId}
          options={feedbackProducts.map((x) => ({
            value: x.id,
            label: `${x.source}（${x.status}）`,
          }))}
        />
        <Button onClick={() => message.success("产品运营数据已刷新")}>
          刷新数据
        </Button>
      </Card>
      <Card className="lineage-card">
        <div className="product-lineage">
          <div>
            <Tag color="orange">来源产品</Tag>
            <b>{item.source}</b>
            <small>
              {item.id} · {item.org}报送
            </small>
          </div>
          <em>经总行查重、优化与标准化建设</em>
          <div>
            <Tag color={item.status === "已上架" ? "red" : "blue"}>
              全行正式产品
            </Tag>
            <b>{item.formal}</b>
            <small>
              {item.formalId} · {item.version}
            </small>
          </div>
        </div>
      </Card>
      <Row gutter={[14, 14]}>
        <Col xs={24} xl={16}>
          <Card
            title="正式产品运营反馈"
            extra={
              <Tag color={item.status === "已上架" ? "green" : "blue"}>
                {item.status}
              </Tag>
            }
          >
            <Row gutter={12} className="feedback-stats">
              <Col span={8}>
                <Statistic title="覆盖机构" value={item.coverage} suffix="家" />
              </Col>
              <Col span={8}>
                <Statistic title="累计运行" value={item.runs} suffix="次" />
              </Col>
              <Col span={8}>
                <Statistic
                  title="用户评分"
                  value={item.rating}
                  precision={1}
                  suffix="分"
                />
              </Col>
            </Row>
            <ReactECharts
              key={item.id}
              style={{ height: 290 }}
              option={{
                tooltip: { trigger: "axis" },
                legend: {
                  data: ["有效使用", "采纳率"],
                  bottom: 0,
                  left: "center",
                  itemGap: 24,
                },
                grid: {
                  left: 50,
                  right: 50,
                  top: 38,
                  bottom: 66,
                  containLabel: true,
                },
                xAxis: {
                  type: "category",
                  data: ["第1周", "第2周", "第3周", "第4周"],
                  axisLabel: { margin: 12 },
                },
                yAxis: [
                  { type: "value" },
                  {
                    type: "value",
                    max: 100,
                    axisLabel: { formatter: "{value}%" },
                  },
                ],
                series: [
                  {
                    name: "有效使用",
                    type: "bar",
                    barMaxWidth: 86,
                    data: item.trend,
                    itemStyle: { color: "#c7000b" },
                  },
                  {
                    name: "采纳率",
                    type: "line",
                    yAxisIndex: 1,
                    data: item.adoption,
                    smooth: true,
                    symbolSize: 8,
                    itemStyle: { color: "#d89614" },
                  },
                ],
              }}
            />
          </Card>
        </Col>
        <Col xs={24} xl={8}>
          <Card title="总行反馈与来源贡献">
            <Timeline
              items={item.feedback.map((x, i) => ({
                color: i === 0 ? "green" : "blue",
                children: x,
              }))}
            />
            <Descriptions
              size="small"
              column={1}
              items={[
                { label: "来源分行", children: item.org },
                { label: "正式产品编号", children: item.formalId },
                { label: "当前版本", children: item.version },
              ]}
            />
            <Button
              block
              type="primary"
              onClick={() => message.success("优化建议已提交至总行运营评价")}
            >
              提交分行优化建议
            </Button>
          </Card>
        </Col>
      </Row>
      <Card title="近期用户反馈" className="recent-feedback">
        <Table
          size="small"
          pagination={false}
          rowKey="id"
          dataSource={[
            {
              id: 1,
              org: "北京分行",
              role: "对公客户经理",
              score: item.rating,
              content: "结果解释清楚，建议增加批量任务分派。",
              status: "已受理",
            },
            {
              id: 2,
              org: "上海分行",
              role: "风险管理人员",
              score: Math.max(4, item.rating - 0.2),
              content: "希望进一步说明指标变化的来源和更新时间。",
              status: "分析中",
            },
            {
              id: 3,
              org: item.org,
              role: "产品运营人员",
              score: item.rating,
              content: "总行标准化版本保留了分行原有核心业务逻辑。",
              status: "已回复",
            },
          ]}
          columns={[
            { title: "反馈机构", dataIndex: "org" },
            { title: "使用岗位", dataIndex: "role" },
            {
              title: "评分",
              dataIndex: "score",
              render: (x: number) => <Rate disabled allowHalf value={x} />,
            },
            { title: "反馈内容", dataIndex: "content" },
            {
              title: "处理状态",
              dataIndex: "status",
              render: (x: string) => <StatusTag status={x} />,
            },
          ]}
        />
      </Card>
    </div>
  );
}

function ReportForm({
  onSubmitted,
  go,
}: {
  onSubmitted: () => void;
  go: (view: View) => void;
}) {
  const [step, setStep] = useState(0);
  const [submitted, setSubmitted] = useState(false);
  const [form] = Form.useForm();
  const [resourceQuery, setResourceQuery] = useState("");
  const [selectedResources, setSelectedResources] = useState<string[]>([
    "EXT-R001",
    "EXT-R003",
    "EXT-R004",
    "EXT-R005",
  ]);
  const [processNodes, setProcessNodes] = useState([
    "多源事件采集",
    "事件抽取与去重",
    "事件标准化",
    "企业主体关联",
    "机会与风险分类",
    "影响评分",
    "客户归属匹配",
    "任务分派与反馈",
  ]);
  const labels = [
    "基本信息",
    "业务任务",
    "数据来源",
    "加工过程",
    "结果交付",
    "运行情况",
    "附件和提交",
  ];
  const next = () => setStep((x) => Math.min(6, x + 1));
  const submit = () => {
    localStorage.setItem("icbc-submitted", "true");
    setSubmitted(true);
    onSubmitted();
    message.success("报送成功，已生成候选产品编号 CP-2026-0801-09");
  };
  if (submitted)
    return (
      <Card className="report-success-card">
        <Result
          status="success"
          title="分行产品报送成功"
          subTitle="某分行半导体产业链客户机会识别已进入总行候选产品区"
          extra={[
            <Button
              key="again"
              onClick={() => {
                setSubmitted(false);
                setStep(0);
              }}
            >
              查看报送内容
            </Button>,
            <Button
              key="next"
              type="primary"
              onClick={() => go("branch-submissions")}
            >
              查看我的报送
            </Button>,
          ]}
        />
        <div className="report-success-detail">
          <Descriptions
            bordered
            column={2}
            items={[
              {
                label: "报送产品名称",
                children: "某分行半导体产业链客户机会识别",
              },
              { label: "候选产品编号", children: <b>CP-2026-0801-09</b> },
              { label: "报送机构", children: "某分行" },
              { label: "报送时间", children: "2026-08-10 16:30" },
              {
                label: "当前状态",
                children: <StatusTag status="待总行受理" />,
              },
              {
                label: "材料完整度",
                children: <b className="red">100%｜校验通过</b>,
              },
            ]}
          />
          <div className="identity-note">
            <SafetyCertificateOutlined />
            <span>
              <b>报送结果说明</b>
              本次报送已形成候选产品及完整建设档案，后续由总行开展完整性校验、存量查重和标准化建设路径判断；报送成功不等同于正式产品上架。
            </span>
          </div>
          <Steps
            className="report-success-flow"
            current={1}
            items={[
              {
                title: "分行报送",
                description: "材料校验通过",
                status: "finish",
              },
              {
                title: "总行受理",
                description: "等待审核与查重",
                status: "process",
              },
              {
                title: "标准化建设",
                description: "形成建设路径",
                status: "wait",
              },
              {
                title: "验证发布",
                description: "上架正式产品",
                status: "wait",
              },
            ]}
          />
        </div>
      </Card>
    );
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            分行产品在线报送 <DemoTag />
          </h1>
          <p>
            提交已组建或试运行的分行产品、运行成效和典型案例，形成总行候选产品
          </p>
        </div>
        <Space>
          <Button onClick={() => message.success("草稿已保存至分行工作台")}>
            保存草稿
          </Button>
          <Button onClick={onSubmitted}>查看报送进度</Button>
        </Space>
      </div>
      <Card className="report-card">
        <Steps
          current={step}
          size="small"
          items={labels.map((title) => ({ title }))}
        />
        <div className="form-body">
          <Form
            form={form}
            layout="vertical"
            initialValues={{
              name: "某分行半导体产业链客户机会识别",
              org: "某分行",
              owner: "陈嘉敏",
              scene: "客户营销",
              role: "对公客户经理、风险管理人员",
              region: ["四川省"],
              uses: 326,
            }}
          >
            {step === 0 && (
              <Row gutter={18}>
                <Col span={12}>
                  <Form.Item
                    label="产品名称"
                    name="name"
                    rules={[{ required: true }]}
                  >
                    <Input />
                  </Form.Item>
                </Col>
                <Col span={12}>
                  <Form.Item
                    label="报送机构"
                    name="org"
                    rules={[{ required: true }]}
                  >
                    <Select
                      showSearch
                      optionFilterProp="label"
                      options={reportingOrganizations.map((value) => ({
                        value,
                        label: value,
                      }))}
                    />
                  </Form.Item>
                </Col>
                <Col span={12}>
                  <Form.Item label="产品负责人" name="owner">
                    <Input />
                  </Form.Item>
                </Col>
                <Col span={12}>
                  <Form.Item label="适用地域" name="region">
                    <Select
                      mode="multiple"
                      showSearch
                      optionFilterProp="label"
                      maxTagCount="responsive"
                      options={domesticRegions.map((value) => ({
                        value,
                        label: value,
                      }))}
                    />
                  </Form.Item>
                </Col>
              </Row>
            )}
            {step === 1 && (
              <Row gutter={18}>
                <Col span={12}>
                  <Form.Item
                    label="业务任务（与外数产品分类一致）"
                    name="scene"
                  >
                    <Select
                      options={productTaskCategories.map((value) => ({
                        value,
                      }))}
                    />
                  </Form.Item>
                </Col>
                <Col span={12}>
                  <Form.Item label="具体任务">
                    <Select
                      defaultValue="产业事件融资机会识别"
                      options={[
                        "产业事件融资机会识别",
                        "重大项目融资线索",
                        "政策影响客户识别",
                        "产业风险事件监测",
                      ].map((value) => ({ value }))}
                    />
                  </Form.Item>
                </Col>
                <Col span={12}>
                  <Form.Item label="使用岗位" name="role">
                    <Input />
                  </Form.Item>
                </Col>
                <Col span={12}>
                  <Form.Item label="嵌入流程节点">
                    <Input defaultValue="营销线索发现、存量客户风险监测与任务分派" />
                  </Form.Item>
                </Col>
                <Col span={24}>
                  <Form.Item label="业务需求说明">
                    <Input.TextArea
                      rows={4}
                      defaultValue="持续采集产业政策、重大项目、招中标、融资、扩产技改、行政处罚及舆情事件，将非结构化信息转化为标准事件并关联企业，识别融资机会与风险信号，推送至对应客户经理和风险人员。"
                    />
                  </Form.Item>
                </Col>
              </Row>
            )}
            {step === 2 && (
              <>
                <Card
                  size="small"
                  title="从外数资源库检索并勾选"
                  extra={
                    <Tag color="green">
                      已选择 {selectedResources.length} 项
                    </Tag>
                  }
                >
                  <Input
                    prefix={<SearchOutlined />}
                    value={resourceQuery}
                    onChange={(e) => setResourceQuery(e.target.value)}
                    allowClear
                    placeholder="搜索资源名称、编号、供应商或字段"
                    className="resource-picker-search"
                  />
                  <div className="resource-picker-list">
                    {resources
                      .filter((r: any) =>
                        `${r.id}${r.name}${r.supplier}${r.fields}`
                          .toLowerCase()
                          .includes(resourceQuery.toLowerCase()),
                      )
                      .map((r: any) => (
                        <label
                          key={r.id}
                          className={
                            selectedResources.includes(r.id) ? "selected" : ""
                          }
                        >
                          <Checkbox
                            checked={selectedResources.includes(r.id)}
                            onChange={(e) =>
                              setSelectedResources(
                                e.target.checked
                                  ? [...selectedResources, r.id]
                                  : selectedResources.filter((x) => x !== r.id),
                              )
                            }
                          />
                          <div>
                            <b>{r.name}</b>
                            <span>
                              {r.id} · {r.type} · {r.supplier}
                            </span>
                            <small>{r.fields}</small>
                          </div>
                          <Tag color={r.auth === "已授权" ? "green" : "orange"}>
                            {r.auth}
                          </Tag>
                        </label>
                      ))}
                  </div>
                </Card>
                <Form.Item label="行内数据来源" style={{ marginTop: 16 }}>
                  <Checkbox.Group
                    options={[
                      "对公客户主数据",
                      "存量授信客户名单",
                      "客户经理管户关系",
                    ]}
                    defaultValue={["对公客户主数据", "客户经理管户关系"]}
                  />
                </Form.Item>
              </>
            )}
            {step === 3 && (
              <>
                <div className="process-model-head">
                  <div>
                    <b>加工链路建模</b>
                    <span>
                      系统依据业务说明和已选资源生成建议步骤；报送人员负责确认、调整和补充，无需从空白文本开始。
                    </span>
                  </div>
                  <Button
                    icon={<RobotOutlined />}
                    onClick={() =>
                      message.success("已根据需求与资源重新生成加工链路建议")
                    }
                  >
                    重新智能识别
                  </Button>
                </div>
                <div className="process-node-editor">
                  {processNodes.map((x, i) => (
                    <React.Fragment key={`${x}-${i}`}>
                      <div>
                        <span>{i + 1}</span>
                        <Input
                          value={x}
                          onChange={(e) =>
                            setProcessNodes(
                              processNodes.map((n, j) =>
                                j === i ? e.target.value : n,
                              ),
                            )
                          }
                        />
                        <Select
                          defaultValue={
                            i < 3
                              ? "标准化处理"
                              : i === 3
                                ? "能力模块"
                                : i === 4
                                  ? "规则判断"
                                  : i === 5
                                    ? "指标加工"
                                    : "输出处理"
                          }
                          options={[
                            "标准化处理",
                            "能力模块",
                            "指标加工",
                            "风险核验",
                            "规则判断",
                            "人工复核",
                            "输出处理",
                          ].map((value) => ({ value }))}
                        />
                        <Button
                          type="text"
                          danger
                          onClick={() =>
                            setProcessNodes(
                              processNodes.filter((_, j) => j !== i),
                            )
                          }
                        >
                          删除
                        </Button>
                      </div>
                      {i < processNodes.length - 1 && <em>↓</em>}
                    </React.Fragment>
                  ))}
                </div>
                <Button
                  block
                  icon={<PlusOutlined />}
                  onClick={() =>
                    setProcessNodes([...processNodes, "请输入新增加工步骤"])
                  }
                >
                  补充自定义步骤
                </Button>
                <Form.Item label="指标和规则" style={{ marginTop: 16 }}>
                  <Input.TextArea
                    rows={3}
                    defaultValue="机会事件包括重大项目中标、扩产技改、政策支持与再融资需求；风险事件包括重大处罚、司法执行和高风险舆情；按事件时效、金额、可信度、客户关系和风险等级综合评分。"
                  />
                </Form.Item>
              </>
            )}
            {step === 4 && (
              <>
                <Row gutter={18}>
                  <Col span={12}>
                    <Form.Item label="交付形式">
                      <Select
                        mode="multiple"
                        defaultValue={["企业名单", "行动提示"]}
                        options={[
                          "企业名单",
                          "企业画像",
                          "指标看板",
                          "风险预警",
                          "专题报告",
                          "API服务",
                          "Excel文件",
                          "任务推送",
                        ].map((value) => ({ value }))}
                      />
                    </Form.Item>
                  </Col>
                  <Col span={12}>
                    <Form.Item label="交付渠道">
                      <Select
                        mode="multiple"
                        defaultValue={[
                          "平台在线查看",
                          "客户经理任务池",
                          "Excel下载",
                        ]}
                        options={[
                          "平台在线查看",
                          "客户经理任务池",
                          "Excel下载",
                          "API推送",
                          "消息提醒",
                          "邮件报表",
                        ].map((value) => ({ value }))}
                      />
                    </Form.Item>
                  </Col>
                  <Col span={8}>
                    <Form.Item label="更新频率">
                      <Select
                        defaultValue="每日"
                        options={[
                          "实时",
                          "每日",
                          "每周",
                          "每月",
                          "按需运行",
                        ].map((value) => ({ value }))}
                      />
                    </Form.Item>
                  </Col>
                  <Col span={8}>
                    <Form.Item label="结果粒度">
                      <Select
                        defaultValue="企业级明细"
                        options={[
                          "企业级明细",
                          "集团级汇总",
                          "地区级汇总",
                          "事件级明细",
                        ].map((value) => ({ value }))}
                      />
                    </Form.Item>
                  </Col>
                  <Col span={8}>
                    <Form.Item label="结果保留期限">
                      <Select
                        defaultValue="12个月"
                        options={["3个月", "6个月", "12个月", "长期留存"].map(
                          (value) => ({ value }),
                        )}
                      />
                    </Form.Item>
                  </Col>
                  <Col span={24}>
                    <Form.Item label="后续业务动作">
                      <Input defaultValue="推送客户经理任务池，记录触达、采纳与营销转化结果" />
                    </Form.Item>
                  </Col>
                </Row>
              </>
            )}
            {step === 5 && (
              <Row gutter={18}>
                <Col span={8}>
                  <Form.Item label="累计使用次数" name="uses">
                    <InputNumber style={{ width: "100%" }} />
                  </Form.Item>
                </Col>
                <Col span={8}>
                  <Form.Item label="运行状态">
                    <Select
                      defaultValue="试运行"
                      options={[{ value: "试运行" }, { value: "正常运行" }]}
                    />
                  </Form.Item>
                </Col>
                <Col span={8}>
                  <Form.Item label="维护主体">
                    <Input defaultValue="某分行数据管理部、公司金融部" />
                  </Form.Item>
                </Col>
                <Col span={24}>
                  <Form.Item label="应用成效">
                    <Input.TextArea
                      rows={4}
                      defaultValue="试运行三个月归集产业事件2.8万条，关联企业6,420家，形成有效融资线索326条、风险提示47条；客户经理已核验线索214条，其中68条进入融资方案沟通。"
                    />
                  </Form.Item>
                </Col>
              </Row>
            )}
            {step === 6 && (
              <>
                <Upload.Dragger>
                  <p className="ant-upload-drag-icon">
                    <CloudUploadOutlined />
                  </p>
                  <p>点击或拖拽上传产品说明、运行截图和使用案例</p>
                  <p className="muted">
                    已准备《某分行半导体产业链客户机会识别产品说明书.pdf》《试运行成效明细.xlsx》（演示文件）
                  </p>
                </Upload.Dragger>
                <Card size="small" className="submit-summary">
                  <b>提交确认</b>
                  <span>已完成 31 项必填字段</span>
                  <span>完整度 100%</span>
                  <StatusTag status="校验通过" />
                </Card>
              </>
            )}
          </Form>
        </div>
        <div className="form-footer">
          <Button disabled={step === 0} onClick={() => setStep((x) => x - 1)}>
            上一步
          </Button>
          <span>第 {step + 1} 步，共 7 步</span>
          {step < 6 ? (
            <Button type="primary" onClick={next}>
              保存并继续
            </Button>
          ) : (
            <Button type="primary" onClick={submit}>
              提交报送
            </Button>
          )}
        </div>
      </Card>
    </div>
  );
}

function CandidateZone({
  goBuild,
  submitted,
  productName,
  onRename,
}: {
  goBuild: () => void;
  submitted: boolean;
  productName: string;
  onRename: (name: string) => void;
}) {
  const [accepted, setAccepted] = useState(
    () =>
      typeof window !== "undefined" &&
      localStorage.getItem("icbc-accepted") === "true",
  );
  const [open, setOpen] = useState(false);
  const [selected, setSelected] = useState<any>(candidates[0]);
  const [editingName, setEditingName] = useState(false);
  const [draftName, setDraftName] = useState(productName);
  const [meta, setMeta] = useState(() => ({
    category: "产业研究＋客户营销＋风险监测",
    scope: "拟全行适用",
    owner: "总行数据管理部、公司金融部、风险管理部",
    position: "产业链监测—事件影响分析—客户经营—授信与贷后协同",
    deliverable:
      "产业链经营驾驶舱、事件影响图、链上客户分层、融资机会包、风险传导路径及协同任务",
  }));
  const [query, setQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("全部状态");
  const [orgFilter, setOrgFilter] = useState("全部分行");
  const [sceneFilter, setSceneFilter] = useState("全部场景");
  const [maturityFilter, setMaturityFilter] = useState("全部成熟度");
  const [scopeFilter, setScopeFilter] = useState("全部范围");
  const [verifyFilter, setVerifyFilter] = useState("全部核验状态");
  const [advancedOpen, setAdvancedOpen] = useState(true);
  const [detailTab, setDetailTab] = useState("overview");
  const [previewRan, setPreviewRan] = useState(false);
  const candidateMeta: any = {
    "CP-2026-0801-09": {
      scene: "客户营销＋风险监测",
      type: "场景型产品",
      maturity: "试运行",
      scope: "拟全行",
      verify: "待核验",
      date: "2026-08-01",
      region: "西部",
    },
    "CP-2026-0731-01": {
      scene: "客户营销",
      type: "场景型产品",
      maturity: "试运行",
      scope: "地方适用",
      verify: "核验中",
      date: "2026-07-31",
      region: "长三角",
    },
    "CP-2026-0728-06": {
      scene: "客户营销",
      type: "功能型产品",
      maturity: "稳定运行",
      scope: "区域适用",
      verify: "已核验",
      date: "2026-07-28",
      region: "长三角",
    },
    "CP-2026-0724-03": {
      scene: "风险监测",
      type: "场景型产品",
      maturity: "试运行",
      scope: "拟全行",
      verify: "待核验",
      date: "2026-07-24",
      region: "粤港澳",
    },
    "CP-2026-0721-08": {
      scene: "客户营销",
      type: "场景型产品",
      maturity: "稳定运行",
      scope: "拟全行",
      verify: "已核验",
      date: "2026-07-21",
      region: "环渤海",
    },
    "CP-2026-0718-05": {
      scene: "产业研究",
      type: "功能型产品",
      maturity: "稳定运行",
      scope: "区域适用",
      verify: "核验中",
      date: "2026-07-18",
      region: "长三角",
    },
    "CP-2026-0715-11": {
      scene: "经营分析",
      type: "功能型产品",
      maturity: "试运行",
      scope: "地方适用",
      verify: "待核验",
      date: "2026-07-15",
      region: "东南",
    },
    "CP-2026-0712-02": {
      scene: "客户营销",
      type: "场景型产品",
      maturity: "已验证",
      scope: "拟全行",
      verify: "已核验",
      date: "2026-07-12",
      region: "西部",
    },
  };
  const rows = candidates.map((x, i) => ({
    ...x,
    ...candidateMeta[x.id],
    ...(i === 0
      ? {
          name: productName,
          status: accepted ? "待查重" : submitted ? "待受理" : x.status,
        }
      : {}),
  }));
  const shownRows = rows.filter(
    (x: any) =>
      (statusFilter === "全部状态" || x.status === statusFilter) &&
      (orgFilter === "全部分行" || x.org === orgFilter) &&
      (sceneFilter === "全部场景" || x.scene === sceneFilter) &&
      (maturityFilter === "全部成熟度" || x.maturity === maturityFilter) &&
      (scopeFilter === "全部范围" || x.scope === scopeFilter) &&
      (verifyFilter === "全部核验状态" || x.verify === verifyFilter) &&
      `${x.name}${x.id}${x.org}${x.owner}`
        .toLowerCase()
        .includes(query.trim().toLowerCase()),
  );
  const resetFilters = () => {
    setQuery("");
    setStatusFilter("全部状态");
    setOrgFilter("全部分行");
    setSceneFilter("全部场景");
    setMaturityFilter("全部成熟度");
    setScopeFilter("全部范围");
    setVerifyFilter("全部核验状态");
  };
  const countBy = (key: string, value: string) =>
    shownRows.filter((x: any) => x[key] === value).length;
  const sceneData = ["客户营销", "风险监测", "产业研究", "经营分析"].map(
    (x) =>
      countBy("scene", x) +
      shownRows.filter(
        (r: any) => r.scene.includes("＋") && r.scene.includes(x),
      ).length,
  );
  const details: any = {
    "CP-2026-0801-09": {
      scene: "产业研究＋客户营销＋风险监测—先进制造产业链经营",
      run: "试运行",
      uses: "326 次",
      desc: "围绕先进制造产业链，融合政策项目、招投标、产业链关系、价格物流、信用评级、工商司法与舆情数据，将产业变化转化为链上客户机会、风险传导、融资方案和协同任务。",
      similar: "产业事件驱动的融资机会识别",
      note: "报送产品已从单点事件线索扩展到产业链影响分析和多岗位协同，但产业链口径、跨区域适配、模型证据与行内客户映射仍需总行标准化。",
    },
    "CP-2026-0731-01": {
      scene: "客户营销—目标客户发现",
      run: "试运行",
      uses: "186 次",
      desc: "整合工商登记、专精特新名单、招投标与司法风险数据，面向客户经理形成浙江辖内优质制造业企业营销名单。",
      similar: "全行专精特新企业营销名单",
      note: "业务任务、使用岗位与核心加工流程高度一致；地域参数及地方名单来源存在差异。",
    },
    "CP-2026-0728-06": {
      scene: "客户营销—招投标商机发现",
      run: "正常运行",
      uses: "94 次",
      desc: "持续汇集江苏地区招标、中标及重点项目公告，识别近期经营活跃且具有融资需求的企业，形成客户拓展线索。",
      similar: "招投标客户发现",
      note: "可直接复用招投标活跃度计算与名单输出能力；需补充江苏区域项目分类和营销阈值。",
    },
    "CP-2026-0724-03": {
      scene: "风险监测—供应链风险识别",
      run: "试运行",
      uses: "61 次",
      desc: "围绕广东制造业核心企业，整合供应链关系、司法涉诉、经营异常与舆情信息，识别上下游风险传导和集中度变化。",
      similar: "核心企业供应链风险画像",
      note: "关系穿透和风险核验能力可复用；供应链依赖度指标及广东产业链范围需要本地适配。",
    },
  };
  const candidateEvidence: any = {
    "CP-2026-0801-09": {
      period: "2026年5月—7月",
      runs: 326,
      coverage: "4条产业链／6,420家",
      outputs: "373项机会与预警",
      adopt: "65.6%",
      conversion: "31.8%",
      saved: "约31人日/月",
      fresh: "事件触发＋每日",
      funnel: [
        ["归集多源事件", "28,460"],
        ["定位产业链", "4条"],
        ["关联链上企业", "6,420"],
        ["形成经营任务", "373"],
        ["业务采纳", "245"],
      ],
      trend: [42, 58, 71, 89, 112, 128],
      cases: [
        {
          key: 1,
          name: "川南精密制造有限公司",
          signal: "设备更新政策＋扩产技改＋中标增长",
          result: "高机会低风险／设备融资需求约1.2亿元",
          action: "生成设备贷与结算组合建议，分派客户经理与授信协同",
        },
        {
          key: 2,
          name: "西部新能源材料集团",
          signal: "关键材料涨价＋客户订单增长＋外部评级稳定",
          result: "机会较高／毛利承压需核验",
          action: "推送流贷方案，同时生成价格风险核验任务",
        },
        {
          key: 3,
          name: "蜀源工程装备股份有限公司",
          signal: "链主新增重大项目＋企业处罚＋负面舆情",
          result: "订单机会与合规风险并存",
          action: "营销暂缓，风险人员先行复核并回填处置结论",
        },
      ],
    },
    "CP-2026-0731-01": {
      period: "2026年4月—7月",
      runs: 186,
      coverage: "3,850家",
      outputs: "542家",
      adopt: "72.0%",
      conversion: "18.6%",
      saved: "约11人日/月",
      fresh: "每周",
      funnel: [
        ["候选企业", "3,850"],
        ["资质匹配", "1,126"],
        ["风险核验", "904"],
        ["推荐名单", "542"],
        ["已触达", "390"],
      ],
      trend: [61, 74, 82, 96, 113, 118],
      cases: [
        {
          key: 1,
          name: "杭州智控科技有限公司",
          signal: "小巨人＋中标活跃",
          result: "优先营销",
          action: "已建立客户联系",
        },
        {
          key: 2,
          name: "宁波精工新材股份有限公司",
          signal: "省级专精特新＋扩产",
          result: "重点关注",
          action: "纳入月度营销计划",
        },
      ],
    },
    "CP-2026-0728-06": {
      period: "2026年3月—7月",
      runs: 94,
      coverage: "8,310条公告",
      outputs: "286条",
      adopt: "61.5%",
      conversion: "24.1%",
      saved: "约9人日/月",
      fresh: "每日",
      funnel: [
        ["招投标公告", "8,310"],
        ["主体匹配", "5,204"],
        ["融资信号", "462"],
        ["有效线索", "286"],
        ["跟进商机", "176"],
      ],
      trend: [35, 46, 52, 63, 71, 82],
      cases: [
        {
          key: 1,
          name: "苏州城建工程有限公司",
          signal: "中标3.2亿元",
          result: "流动资金机会",
          action: "已推送属地支行",
        },
        {
          key: 2,
          name: "常州智能装备集团",
          signal: "连续中标＋履约保函",
          result: "综合融资机会",
          action: "客户经理核验中",
        },
      ],
    },
    "CP-2026-0724-03": {
      period: "2026年4月—7月",
      runs: 61,
      coverage: "1,980家",
      outputs: "137条",
      adopt: "83.2%",
      conversion: "—",
      saved: "约7人日/月",
      fresh: "准实时",
      funnel: [
        ["监测企业", "1,980"],
        ["关系主体", "8,462"],
        ["风险事件", "394"],
        ["有效预警", "137"],
        ["已处置", "114"],
      ],
      trend: [18, 22, 25, 31, 28, 34],
      cases: [
        {
          key: 1,
          name: "粤东电子零部件有限公司",
          signal: "核心客户涉诉",
          result: "传导风险 中",
          action: "已补充风险排查",
        },
        {
          key: 2,
          name: "南粤汽车部件集团",
          signal: "供应商集中度上升",
          result: "集中度预警",
          action: "纳入贷后检查",
        },
      ],
    },
  };
  const show = (r: any) => {
    setSelected(r);
    setDetailTab("overview");
    setPreviewRan(false);
    setOpen(true);
  };
  const d = details[selected.id] || details["CP-2026-0801-09"];
  const ev =
    candidateEvidence[selected.id] || candidateEvidence["CP-2026-0801-09"];
  const previewOutput =
    selected.id === "CP-2026-0731-01"
      ? [
          "目标企业 542 家",
          "高优先级 126 家",
          "风险排除 222 家",
          "已生成营销名单",
        ]
      : selected.id === "CP-2026-0728-06"
        ? [
            "有效商机 286 条",
            "融资需求强 74 条",
            "今日新增 19 条",
            "已生成客户任务",
          ]
        : selected.id === "CP-2026-0724-03"
          ? [
              "风险预警 137 条",
              "高风险 18 条",
              "传导路径 243 条",
              "已生成处置任务",
            ]
          : [
              "覆盖4条产业链",
              "经营机会326项",
              "风险传导47条",
              "已生成营销/授信/风险协同任务",
            ];
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            候选产品区 <DemoTag />
          </h1>
          <p>掌握各分行产品建设情况，集中开展受理、比较、标准化与推广判断</p>
        </div>
        <Space>
          <Button onClick={() => setAdvancedOpen(!advancedOpen)}>
            {advancedOpen ? "收起建设看板" : "展开建设看板"}
          </Button>
          <Button
            type="primary"
            icon={<DownloadOutlined />}
            onClick={() =>
              message.success(
                `已导出当前筛选的 ${shownRows.length} 条候选产品台账（演示）`,
              )
            }
          >
            导出候选产品台账
          </Button>
        </Space>
      </div>
      {advancedOpen && (
        <div className="candidate-dashboard">
          <Row gutter={[12, 12]}>
            {[
              ["候选产品", shownRows.length, "当前筛选"],
              [
                "报送分行",
                new Set(shownRows.map((x: any) => x.org)).size,
                "涉及机构",
              ],
              ["拟全行推广", countBy("scope", "拟全行"), "推广储备"],
              ["已核验成效", countBy("verify", "已核验"), "可进入评估"],
              [
                "建设及验证中",
                shownRows.filter((x: any) =>
                  ["建设中", "待验证"].includes(x.status),
                ).length,
                "重点跟踪",
              ],
              [
                "平均相似度",
                shownRows.length
                  ? Math.round(
                      shownRows.reduce((s: number, x: any) => s + x.score, 0) /
                        shownRows.length,
                    )
                  : 0,
                "存量复用参考",
              ],
            ].map((x: any) => (
              <Col xs={12} md={8} xl={4} key={x[0]}>
                <Card size="small" className="candidate-stat">
                  <Statistic
                    title={x[0]}
                    value={x[1]}
                    suffix={x[0] === "平均相似度" ? "%" : undefined}
                  />
                  <small>{x[2]}</small>
                </Card>
              </Col>
            ))}
          </Row>
          <Row gutter={12} className="candidate-charts">
            <Col xs={24} xl={12}>
              <Card
                size="small"
                title="候选产品业务场景分布"
                extra={<span>点击图例可查看</span>}
              >
                <ReactECharts
                  style={{ height: 210 }}
                  option={{
                    grid: { left: 48, right: 18, top: 20, bottom: 30 },
                    xAxis: {
                      type: "category",
                      data: ["客户营销", "风险监测", "产业研究", "经营分析"],
                    },
                    yAxis: { type: "value", minInterval: 1 },
                    tooltip: { trigger: "axis" },
                    series: [
                      {
                        type: "bar",
                        barWidth: 28,
                        data: sceneData,
                        itemStyle: {
                          color: "#c7000b",
                          borderRadius: [4, 4, 0, 0],
                        },
                      },
                    ],
                  }}
                />
              </Card>
            </Col>
            <Col xs={24} xl={12}>
              <Card
                size="small"
                title="报送产品建设阶段"
                extra={<span>随筛选联动</span>}
              >
                <ReactECharts
                  style={{ height: 210 }}
                  option={{
                    tooltip: { trigger: "item" },
                    legend: { bottom: 0 },
                    series: [
                      {
                        type: "pie",
                        radius: [42, 72],
                        center: ["50%", "45%"],
                        data: [
                          "待受理",
                          "待查重",
                          "待方案设计",
                          "建设中",
                          "待验证",
                        ]
                          .map((name, i) => ({
                            name,
                            value: countBy("status", name),
                            itemStyle: {
                              color: [
                                "#9ca3af",
                                "#d99000",
                                "#3b82f6",
                                "#c7000b",
                                "#10b981",
                              ][i],
                            },
                          }))
                          .filter((x) => x.value),
                      },
                    ],
                  }}
                />
              </Card>
            </Col>
          </Row>
        </div>
      )}
      <Card className="candidate-ledger">
        <div className="candidate-filter-head">
          <div>
            <b>候选产品台账</b>
            <span>统计图与台账共用筛选条件</span>
          </div>
          <Button type="link" onClick={resetFilters}>
            重置全部筛选
          </Button>
        </div>
        <div className="candidate-advanced-filter">
          <Input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            prefix={<SearchOutlined />}
            placeholder="搜索产品、编号、分行或负责人"
            allowClear
          />
          <Select
            value={orgFilter}
            onChange={setOrgFilter}
            options={[
              "全部分行",
              ...Array.from(new Set(rows.map((x: any) => x.org))),
            ].map((value) => ({ value }))}
          />
          <Select
            value={sceneFilter}
            onChange={setSceneFilter}
            options={[
              "全部场景",
              ...Array.from(new Set(rows.map((x: any) => x.scene))),
            ].map((value) => ({ value }))}
          />
          <Select
            value={statusFilter}
            onChange={setStatusFilter}
            options={[
              "全部状态",
              ...Array.from(new Set(rows.map((x: any) => x.status))),
            ].map((value) => ({ value }))}
          />
          <Select
            value={maturityFilter}
            onChange={setMaturityFilter}
            options={[
              "全部成熟度",
              ...Array.from(new Set(rows.map((x: any) => x.maturity))),
            ].map((value) => ({ value }))}
          />
          <Select
            value={scopeFilter}
            onChange={setScopeFilter}
            options={[
              "全部范围",
              ...Array.from(new Set(rows.map((x: any) => x.scope))),
            ].map((value) => ({ value }))}
          />
          <Select
            value={verifyFilter}
            onChange={setVerifyFilter}
            options={[
              "全部核验状态",
              ...Array.from(new Set(rows.map((x: any) => x.verify))),
            ].map((value) => ({ value }))}
          />
        </div>
        <div className="filter-result">
          <span>
            当前匹配 <b>{shownRows.length}</b> 项，涉及{" "}
            <b>{new Set(shownRows.map((x: any) => x.org)).size}</b> 家分行
          </span>
          {shownRows.length !== rows.length && (
            <Tag
              closable
              onClose={(e) => {
                e.preventDefault();
                resetFilters();
              }}
            >
              已启用高级筛选
            </Tag>
          )}
        </div>
        <Table
          scroll={{ x: 1420 }}
          pagination={{ pageSize: 8, showSizeChanger: false }}
          dataSource={shownRows}
          rowKey="id"
          locale={{
            emptyText: (
              <Empty description="未找到符合条件的候选产品">
                <Button type="primary" onClick={resetFilters}>
                  清空筛选
                </Button>
              </Empty>
            ),
          }}
          columns={[
            {
              title: "候选产品",
              dataIndex: "name",
              width: 260,
              fixed: "left",
              render: (t: string, r: any) => (
                <div>
                  <a onClick={() => show(r)}>{t}</a>
                  <small className="cell-sub">
                    {r.id} · 报送于 {r.date}
                  </small>
                </div>
              ),
            },
            {
              title: "报送分行",
              dataIndex: "org",
              width: 110,
              render: (x: string) => <b>{x}</b>,
            },
            { title: "业务场景", dataIndex: "scene", width: 145 },
            { title: "产品形态", dataIndex: "type", width: 110 },
            {
              title: "成熟度",
              dataIndex: "maturity",
              width: 100,
              render: (x: string) => (
                <Tag
                  color={x === "稳定运行" || x === "已验证" ? "green" : "blue"}
                >
                  {x}
                </Tag>
              ),
            },
            { title: "拟适用范围", dataIndex: "scope", width: 110 },
            {
              title: "成效核验",
              dataIndex: "verify",
              width: 105,
              render: (x: string) => (
                <Tag
                  color={
                    x === "已核验"
                      ? "green"
                      : x === "核验中"
                        ? "blue"
                        : "orange"
                  }
                >
                  {x}
                </Tag>
              ),
            },
            { title: "负责人", dataIndex: "owner", width: 90 },
            {
              title: "相似度",
              dataIndex: "score",
              width: 130,
              render: (x: number) => (
                <Progress
                  percent={x}
                  size="small"
                  strokeColor={x > 85 ? "#c7000b" : "#d99000"}
                />
              ),
            },
            {
              title: "当前状态",
              dataIndex: "status",
              render: (x: string) => <StatusTag status={x} />,
            },
            {
              title: "操作",
              width: 150,
              fixed: "right",
              render: (_: any, r: any) =>
                r.id === rows[0].id ? (
                  <Space>
                    <Button size="small" onClick={() => show(r)}>
                      查看
                    </Button>
                    {!accepted ? (
                      <Button
                        size="small"
                        type="primary"
                        onClick={() => {
                          setAccepted(true);
                          localStorage.setItem("icbc-accepted", "true");
                          message.success("已受理，产品进入待查重状态");
                        }}
                      >
                        受理
                      </Button>
                    ) : (
                      <Button size="small" type="primary" onClick={goBuild}>
                        建设
                      </Button>
                    )}
                  </Space>
                ) : (
                  <Button size="small" onClick={() => show(r)}>
                    查看详情
                  </Button>
                ),
            },
          ]}
        />
      </Card>
      <Drawer
        open={open}
        width={860}
        title={selected.id === rows[0].id ? productName : selected.name}
        onClose={() => {
          setOpen(false);
          setEditingName(false);
        }}
        extra={
          selected.id === rows[0].id && (
            <Space>
              {editingName ? (
                <>
                  <Button
                    onClick={() => {
                      setEditingName(false);
                      setDraftName(productName);
                    }}
                  >
                    取消
                  </Button>
                  <Button
                    type="primary"
                    disabled={!draftName.trim()}
                    onClick={() => {
                      onRename(draftName.trim());
                      setSelected({ ...selected, name: draftName.trim() });
                      setEditingName(false);
                      message.success(
                        "产品名称已更新，原报送名称已保留在变更记录中",
                      );
                    }}
                  >
                    保存名称
                  </Button>
                </>
              ) : (
                <Button
                  icon={<FormOutlined />}
                  onClick={() => {
                    setDraftName(productName);
                    setEditingName(true);
                  }}
                >
                  编辑名称
                </Button>
              )}
              {!accepted ? (
                <Button
                  type="primary"
                  onClick={() => {
                    setAccepted(true);
                    localStorage.setItem("icbc-accepted", "true");
                    message.success("受理成功，产品进入待查重状态");
                  }}
                >
                  受理产品
                </Button>
              ) : (
                <Button type="primary" onClick={goBuild}>
                  进入产品建设中心
                </Button>
              )}
            </Space>
          )
        }
      >
        {selected.id === rows[0].id && (
          <Card
            size="small"
            style={{ marginBottom: 16 }}
            title="总行标准化方案编辑"
          >
            {editingName ? (
              <Form layout="vertical">
                <Form.Item label="标准产品名称">
                  <Input
                    value={draftName}
                    onChange={(e) => setDraftName(e.target.value)}
                    maxLength={60}
                  />
                </Form.Item>
                <Row gutter={12}>
                  <Col span={12}>
                    <Form.Item label="业务定位">
                      <Input
                        value={meta.category}
                        onChange={(e) =>
                          setMeta({ ...meta, category: e.target.value })
                        }
                      />
                    </Form.Item>
                  </Col>
                  <Col span={12}>
                    <Form.Item label="拟适用层级">
                      <Select
                        value={meta.scope}
                        onChange={(scope) => setMeta({ ...meta, scope })}
                        options={[
                          "分行适用",
                          "区域适用",
                          "限域适用",
                          "拟全行适用",
                        ].map((value) => ({ value }))}
                      />
                    </Form.Item>
                  </Col>
                </Row>
                <Form.Item label="总行维护责任">
                  <Input
                    value={meta.owner}
                    onChange={(e) =>
                      setMeta({ ...meta, owner: e.target.value })
                    }
                  />
                </Form.Item>
                <Form.Item label="嵌入业务流程">
                  <Input
                    value={meta.position}
                    onChange={(e) =>
                      setMeta({ ...meta, position: e.target.value })
                    }
                  />
                </Form.Item>
                <Form.Item label="标准交付与决策结果">
                  <Input.TextArea
                    rows={3}
                    value={meta.deliverable}
                    onChange={(e) =>
                      setMeta({ ...meta, deliverable: e.target.value })
                    }
                  />
                </Form.Item>
              </Form>
            ) : (
              <Descriptions
                size="small"
                column={1}
                items={[
                  { label: "标准产品名称", children: productName },
                  {
                    label: "分行原报送名称",
                    children: "某分行半导体产业链客户机会识别",
                  },
                  { label: "业务定位", children: meta.category },
                  { label: "拟适用层级", children: meta.scope },
                  { label: "维护责任", children: meta.owner },
                  { label: "业务流程", children: meta.position },
                  { label: "标准交付", children: meta.deliverable },
                  {
                    label: "变更记录",
                    children:
                      productName === "产业事件驱动的融资机会与风险识别"
                        ? "尚未调整"
                        : "总行产品管理员调整 · 已留痕",
                  },
                ]}
              />
            )}
          </Card>
        )}
        <Tabs
          activeKey={detailTab}
          onChange={setDetailTab}
          items={[
            {
              key: "overview",
              label: "报送概览",
              children: (
                <>
                  <Descriptions
                    column={2}
                    bordered
                    size="small"
                    items={[
                      { label: "候选产品编号", children: selected.id },
                      { label: "报送机构", children: selected.org },
                      { label: "业务场景", children: d.scene },
                      {
                        label: "运行状态",
                        children: <StatusTag status={d.run} />,
                      },
                      { label: "使用次数", children: d.uses },
                      {
                        label: "完整性校验",
                        children: <StatusTag status="通过" />,
                      },
                    ]}
                  />
                  <h3 className="drawer-title">产品说明</h3>
                  <p>{d.desc}</p>
                  <h3 className="drawer-title">相似产品</h3>
                  <Card size="small">
                    <b>{d.similar}</b>
                    <Progress percent={selected.score} strokeColor="#c7000b" />
                    <p className="muted">{d.note}</p>
                  </Card>
                  <h3 className="drawer-title">流程记录</h3>
                  <Timeline
                    items={[
                      {
                        color: "green",
                        children: `产品报送提交 · ${selected.owner.replace(/\s/g, "")}`,
                      },
                      { color: "green", children: "完整性校验通过 · 系统" },
                      {
                        color:
                          selected.id === rows[0].id && !accepted
                            ? "blue"
                            : "green",
                        children:
                          selected.id === rows[0].id && !accepted
                            ? "等待总行产品管理员受理"
                            : `当前状态：${selected.status}`,
                      },
                    ]}
                  />
                </>
              ),
            },
            {
              key: "preview",
              label: "功能预览",
              children: (
                <div className="candidate-preview">
                  <div className="candidate-preview-head">
                    <div>
                      <Tag color="red">分行业务端预览</Tag>
                      <h2>{selected.name}</h2>
                      <p>{d.desc}</p>
                    </div>
                    <Button
                      type="primary"
                      icon={<RobotOutlined />}
                      loading={previewRan === null}
                      onClick={() => {
                        setPreviewRan(true);
                        message.success("演示运行完成，已生成模拟结果");
                      }}
                    >
                      运行演示
                    </Button>
                  </div>
                  <Row gutter={14}>
                    <Col span={9}>
                      <Card size="small" title="运行条件">
                        <Form layout="vertical">
                          <Form.Item label="适用地域">
                            <Select
                              defaultValue={selected.org.replace("分行", "")}
                              options={[
                                { value: selected.org.replace("分行", "") },
                              ]}
                            />
                          </Form.Item>
                          <Form.Item label="观察周期">
                            <Select
                              defaultValue="近90天"
                              options={[
                                { value: "近30天" },
                                { value: "近90天" },
                                { value: "近一年" },
                              ]}
                            />
                          </Form.Item>
                          <Form.Item label="业务对象">
                            <Input
                              defaultValue={
                                selected.id === rows[0].id
                                  ? "先进制造产业链、链上企业及存量客户"
                                  : "辖内企业及存量客户"
                              }
                            />
                          </Form.Item>
                          {selected.id === rows[0].id && (
                            <Form.Item label="已选外数资源">
                              <Select
                                mode="multiple"
                                maxTagCount={3}
                                defaultValue={[
                                  "产业政策与重大项目",
                                  "产业链图谱",
                                  "招投标与订单",
                                  "大宗商品价格",
                                  "物流景气",
                                  "企业信用评级",
                                  "工商司法与舆情",
                                ]}
                                options={[
                                  "产业政策与重大项目",
                                  "产业链图谱",
                                  "招投标与订单",
                                  "大宗商品价格",
                                  "物流景气",
                                  "企业信用评级",
                                  "工商司法与舆情",
                                ].map((value) => ({ value }))}
                              />
                            </Form.Item>
                          )}
                          <Form.Item label="输出类型">
                            <Checkbox.Group
                              defaultValue={["驾驶舱", "判断", "任务"]}
                              options={
                                selected.id === rows[0].id
                                  ? [
                                      "驾驶舱",
                                      "影响图",
                                      "机会包",
                                      "风险路径",
                                      "任务",
                                    ]
                                  : ["名单", "判断", "任务"]
                              }
                            />
                          </Form.Item>
                        </Form>
                      </Card>
                    </Col>
                    <Col span={15}>
                      <Card size="small" title="输出预览" extra={<DemoTag />}>
                        {!previewRan ? (
                          <Empty
                            image={Empty.PRESENTED_IMAGE_SIMPLE}
                            description="设置条件后运行，预览分行产品的输入、处理和输出"
                          />
                        ) : (
                          <>
                            <div className="preview-output-grid">
                              {previewOutput.map((x: string, i: number) => (
                                <div key={x}>
                                  <span>
                                    {
                                      [
                                        "结果规模",
                                        "重点结果",
                                        "补充判断",
                                        "后续动作",
                                      ][i]
                                    }
                                  </span>
                                  <b>{x}</b>
                                </div>
                              ))}
                            </div>
                            <Table
                              size="small"
                              pagination={false}
                              dataSource={ev.cases.slice(0, 2)}
                              columns={[
                                { title: "企业", dataIndex: "name" },
                                { title: "识别信号", dataIndex: "signal" },
                                { title: "产品判断", dataIndex: "result" },
                                { title: "业务动作", dataIndex: "action" },
                              ]}
                            />
                          </>
                        )}
                      </Card>
                    </Col>
                  </Row>
                  <Card size="small" className="preview-chain">
                    <b>分行报送的功能链</b>
                    <div>
                      {[
                        "数据输入",
                        "主体关联",
                        "特征加工",
                        "规则/模型判断",
                        "结果解释",
                        "名单或任务输出",
                      ].map((x: string, i: number) => (
                        <React.Fragment key={x}>
                          <span>{x}</span>
                          {i < 5 && <em>→</em>}
                        </React.Fragment>
                      ))}
                    </div>
                    <p>
                      功能预览用于总行受理前理解产品，不代表已完成全行标准化改造；资源授权、模型口径和适用范围仍需后续核验。
                    </p>
                  </Card>
                </div>
              ),
            },
            {
              key: "effect",
              label: "运行成效",
              children: (
                <div className="candidate-effect">
                  <div className="effect-alert">
                    <SafetyCertificateOutlined />
                    <div>
                      <b>以下为分行填报的运行效果，尚待总行核验</b>
                      <span>
                        统计周期：{ev.period}　更新频率：{ev.fresh}
                        　可结合运行明细和佐证材料复核
                      </span>
                    </div>
                    <Button
                      size="small"
                      onClick={() => message.success("已发起成效数据核验任务")}
                    >
                      发起核验
                    </Button>
                  </div>
                  <div className="effect-metrics">
                    {[
                      ["累计运行", ev.runs, "次"],
                      ["业务覆盖", ev.coverage, ""],
                      ["有效输出", ev.outputs, ""],
                      ["业务采纳率", ev.adopt, ""],
                      ["转化/处置率", ev.conversion, ""],
                      ["估算节省工时", ev.saved, ""],
                    ].map((x: any) => (
                      <Card size="small" key={x[0]}>
                        <Statistic title={x[0]} value={x[1]} suffix={x[2]} />
                      </Card>
                    ))}
                  </div>
                  <Card
                    size="small"
                    title="运行漏斗与月度有效输出"
                    extra={<Tag color="orange">分行填报口径</Tag>}
                  >
                    <Row gutter={18}>
                      <Col span={14}>
                        <div className="effect-funnel">
                          {ev.funnel.map((x: any, i: number) => (
                            <React.Fragment key={x[0]}>
                              <div>
                                <span>{x[0]}</span>
                                <b>{x[1]}</b>
                              </div>
                              {i < ev.funnel.length - 1 && <em>→</em>}
                            </React.Fragment>
                          ))}
                        </div>
                      </Col>
                      <Col span={10}>
                        <ReactECharts
                          style={{ height: 180 }}
                          option={{
                            grid: { left: 38, right: 12, top: 18, bottom: 28 },
                            xAxis: {
                              type: "category",
                              data: ["2月", "3月", "4月", "5月", "6月", "7月"],
                            },
                            yAxis: { type: "value" },
                            series: [
                              {
                                type: "line",
                                smooth: true,
                                data: ev.trend,
                                lineStyle: { color: "#c7000b" },
                                itemStyle: { color: "#c7000b" },
                                areaStyle: { color: "rgba(199,0,11,.08)" },
                              },
                            ],
                            tooltip: { trigger: "axis" },
                          }}
                        />
                      </Col>
                    </Row>
                  </Card>
                  <Card
                    size="small"
                    title="典型运行案例"
                    extra={
                      <Button
                        size="small"
                        icon={<DownloadOutlined />}
                        onClick={() =>
                          message.success("运行明细已导出（演示）")
                        }
                      >
                        导出明细
                      </Button>
                    }
                  >
                    <Table
                      size="small"
                      pagination={false}
                      dataSource={ev.cases}
                      columns={[
                        { title: "业务对象", dataIndex: "name" },
                        { title: "触发信号", dataIndex: "signal" },
                        {
                          title: "产品输出",
                          dataIndex: "result",
                          render: (x: string) => (
                            <Tag color={x.includes("风险") ? "orange" : "red"}>
                              {x}
                            </Tag>
                          ),
                        },
                        { title: "实际业务动作", dataIndex: "action" },
                      ]}
                    />
                  </Card>
                  <Card size="small" title="成效口径与佐证材料">
                    <Descriptions
                      size="small"
                      column={2}
                      items={[
                        {
                          label: "有效输出",
                          children:
                            "经去重且达到分行产品阈值的名单、预警或机会线索",
                        },
                        {
                          label: "采纳率",
                          children:
                            "已被业务人员领取、查看或进入处置的有效输出占比",
                        },
                        {
                          label: "转化/处置率",
                          children:
                            "进入营销沟通或完成风险处置的结果占已核验结果比例",
                        },
                        {
                          label: "节省工时",
                          children: "分行按原人工流程与产品运行耗时差异估算",
                        },
                        {
                          label: "已附材料",
                          children:
                            "运行成效明细.xlsx、典型案例说明.pdf、业务部门评价表.pdf",
                        },
                        {
                          label: "总行核验状态",
                          children: <StatusTag status="待审核" />,
                        },
                      ]}
                    />
                  </Card>
                </div>
              ),
            },
            {
              key: "assessment",
              label: "总行初评",
              children: (
                <Card title="受理前初步判断">
                  <Descriptions
                    bordered
                    column={2}
                    items={[
                      {
                        label: "业务价值",
                        children:
                          "可将分散事件转为机会、风险判断和业务任务，具备跨机构复用潜力",
                      },
                      {
                        label: "成效可信度",
                        children: <StatusTag status="待审核" />,
                      },
                      {
                        label: "标准化重点",
                        children:
                          "统一事件口径、模型阈值、输出解释及任务反馈字段",
                      },
                      { label: "适用范围建议", children: meta.scope },
                      {
                        label: "拟采用路径",
                        children: "能力组合＋标准产品升级",
                      },
                      {
                        label: "下一步",
                        children: "受理后开展产品查重、资源核验和成效复核",
                      },
                    ]}
                  />
                  <div className="pane-actions">
                    <Button onClick={() => setDetailTab("effect")}>
                      返回查看成效
                    </Button>
                    <Button
                      type="primary"
                      onClick={() => {
                        if (!accepted) {
                          setAccepted(true);
                          localStorage.setItem("icbc-accepted", "true");
                        }
                        message.success("已记录总行初评，候选产品进入待查重");
                      }}
                    >
                      确认初评并受理
                    </Button>
                  </div>
                </Card>
              ),
            },
          ]}
        />
      </Drawer>
    </div>
  );
}

function Requirements({ next, profile }: { next: () => void; profile: any }) {
  return (
    <div className="build-pane">
      <div className="five-blocks">
        {[
          ["业务任务", profile.task, `${profile.role}｜${profile.process}`],
          ["数据输入", profile.inputs, "与报送的数据来源保持一致"],
          ["加工处理", profile.processing, "由报送链路自动带入，可复核调整"],
          ["结果交付", profile.output, "与报送交付形式保持一致"],
          ["后续动作", profile.action, "形成业务反馈闭环"],
        ].map((x, i) => (
          <Card key={x[0]} className="require-card">
            <span>0{i + 1}</span>
            <h3>{x[0]}</h3>
            <b>{x[1]}</b>
            <p>{x[2]}</p>
            <Button type="link">编辑</Button>
          </Card>
        ))}
      </div>
      <Card className="baseline-card">
        <SectionTitle
          title="结构化需求基线"
          sub="系统已从报送材料自动提取，可在确认前修改"
        />
        <Descriptions
          bordered
          size="small"
          column={3}
          items={[
            { label: "业务分类", children: profile.category },
            { label: "使用岗位", children: profile.role },
            { label: "流程节点", children: profile.process },
            { label: "数据输入", children: profile.inputs },
            { label: "结果形式", children: profile.output },
            { label: "适用地域", children: profile.region },
          ]}
        />
        <div className="pane-actions">
          <Button onClick={() => message.success("需求基线草稿已保存")}>
            保存草稿
          </Button>
          <Button
            type="primary"
            onClick={() => {
              message.success("需求基线已确认");
              next();
            }}
          >
            确认需求基线
          </Button>
        </div>
      </Card>
    </div>
  );
}

function SceneMapping({ next, profile }: { next: () => void; profile: any }) {
  const eventCase = profile.task.includes("事件");
  const flow = eventCase
    ? [
        "产业事件发现",
        "事件标准化",
        "主体与客户关联",
        "机会/风险判断",
        "任务分派",
        "结果反馈",
      ]
    : ["目标客户发现", "企业初筛", "名单形成", "客户触达", "营销反馈"];
  const [selected, setSelected] = useState(
    eventCase ? "主体与客户关联" : "企业初筛",
  );
  return (
    <div className="build-pane">
      <Row gutter={16}>
        <Col span={8}>
          <Card title="业务场景映射">
            <Form layout="vertical">
              <Form.Item label="业务场景分类">
                <Select
                  defaultValue={profile.category}
                  options={productTaskCategories.map((value) => ({ value }))}
                />
              </Form.Item>
              <Form.Item label="具体业务任务">
                <Input.TextArea
                  rows={3}
                  defaultValue={profile.task}
                  placeholder="请用业务语言描述需要完成的具体任务"
                />
              </Form.Item>
              <Form.Item label="嵌入流程节点">
                <Input defaultValue={profile.process} />
              </Form.Item>
              <Form.Item label="使用岗位">
                <Input defaultValue={profile.role} />
              </Form.Item>
              <Form.Item label="支持层级">
                <Checkbox.Group
                  options={[
                    "数据查询",
                    "指标加工",
                    "规则判断",
                    "行动提示",
                    "业务反馈",
                  ]}
                  defaultValue={[
                    "数据查询",
                    "指标加工",
                    "规则判断",
                    "行动提示",
                  ]}
                />
              </Form.Item>
            </Form>
          </Card>
        </Col>
        <Col span={16}>
          <Card title="产品所在业务流程">
            <div className="business-flow scene-business-flow">
              {flow.map((x, i) => (
                <React.Fragment key={x}>
                  <button
                    className={selected === x ? "selected" : ""}
                    onClick={() => setSelected(x)}
                  >
                    <span>{i + 1}</span>
                    <b>{x}</b>
                    <small>
                      {selected === x ? "产品嵌入节点" : "上下游流程"}
                    </small>
                  </button>
                  {i < flow.length - 1 && <em>→</em>}
                </React.Fragment>
              ))}
            </div>
            <div className="mapping-callout">
              <ApartmentOutlined />
              <div>
                <b>已映射至：公司金融 → 产业事件发现 → 机会与风险双向分派</b>
                <p>
                  产品把公开事件转化为企业级业务信号，并分别进入营销拓展与风险处置流程。
                </p>
              </div>
            </div>
          </Card>
        </Col>
      </Row>
      <div className="pane-actions">
        <Button>保存映射</Button>
        <Button type="primary" onClick={next}>
          确认并开始查重
        </Button>
      </div>
    </div>
  );
}

function Dedupe({ next, go }: { next: () => void; go: (v: View) => void }) {
  const radar = {
    tooltip: {},
    radar: {
      indicator: [
        { name: "业务任务", max: 100 },
        { name: "目标用户", max: 100 },
        { name: "流程节点", max: 100 },
        { name: "数据输入", max: 100 },
        { name: "加工逻辑", max: 100 },
        { name: "主要输出", max: 100 },
        { name: "适用范围", max: 100 },
      ],
      splitArea: { areaStyle: { color: ["#fff", "#fafafa"] } },
    },
    series: [
      {
        type: "radar",
        data: [{ value: [86, 92, 78, 74, 69, 81, 88], name: "相似度" }],
        areaStyle: { color: "rgba(199,0,11,.16)" },
        lineStyle: { color: "#c7000b" },
        itemStyle: { color: "#c7000b" },
      },
    ],
  };
  return (
    <div className="build-pane">
      <Tabs
        defaultActiveKey="product"
        items={[
          {
            key: "product",
            label: "产品查重",
            children: (
              <Row gutter={16}>
                <Col span={9}>
                  <Card title="综合相似度">
                    <div className="similar-score">
                      <strong>76</strong>
                      <span>%</span>
                      <b>部分重合</b>
                    </div>
                    <ReactECharts style={{ height: 300 }} option={radar} />
                  </Card>
                </Col>
                <Col span={15}>
                  <Card title="最相似产品">
                    <div className="similar-product">
                      <div>
                        <Tag color="red">正式产品</Tag>
                        <h2>产业事件驱动的融资机会识别</h2>
                        <p>P-2026-001 · 全行适用 · 当前版本 V2.1</p>
                      </div>
                      <Progress
                        type="circle"
                        percent={76}
                        size={92}
                        strokeColor="#c7000b"
                      />
                    </div>
                    <Table
                      size="small"
                      pagination={false}
                      dataSource={[
                        {
                          key: 1,
                          d: "业务任务",
                          same: "均识别产业事件形成融资机会",
                          diff: "新增风险信号识别与双向任务分派",
                          score: 86,
                        },
                        {
                          key: 2,
                          d: "数据输入",
                          same: "政策、项目、招投标与工商数据一致",
                          diff: "增加处罚、司法、舆情与客户归属",
                          score: 74,
                        },
                        {
                          key: 3,
                          d: "加工逻辑",
                          same: "事件抽取—主体关联—机会评分",
                          diff: "增加风险分类、影响范围与交叉核验",
                          score: 69,
                        },
                        {
                          key: 4,
                          d: "适用范围",
                          same: "对公客户经理与公司金融岗位",
                          diff: "新增风险人员，四川样本拟验证全行母版",
                          score: 88,
                        },
                      ]}
                      columns={[
                        { title: "比对维度", dataIndex: "d" },
                        { title: "相同内容", dataIndex: "same" },
                        { title: "主要差异", dataIndex: "diff" },
                        {
                          title: "相似度",
                          dataIndex: "score",
                          render: (v: number) => (
                            <Progress percent={v} size="small" />
                          ),
                        },
                      ]}
                    />
                    <div className="recommend-box">
                      <CheckCircleFilled />
                      <div>
                        <b>系统建议：组合能力模块并升级为全行产品母版</b>
                        <p>
                          复用既有机会识别链路，新增风险事件分类、事件影响评分和营销/风险双任务分派，完成异地样本验证后形成全行标准产品。
                        </p>
                      </div>
                    </div>
                  </Card>
                </Col>
              </Row>
            ),
          },
          {
            key: "resource",
            label: "资源覆盖与采购缺口",
            children: (
              <>
                <Row gutter={14}>
                  {[
                    ["需求字段", "46", "业务需求拆解"],
                    ["直接覆盖", "31", "现有合同可用"],
                    ["语义映射", "9", "需统一事件口径"],
                    ["采购缺口", "3", "形成授权扩围事项"],
                  ].map((x) => (
                    <Col span={6} key={x[0]}>
                      <MetricCard title={x[0]} value={x[1]} trend={x[2]} />
                    </Col>
                  ))}
                </Row>
                <Card
                  title="字段级覆盖与处置矩阵"
                  extra={<Tag color="green">预计避免重复采购 74%</Tag>}
                >
                  <Table
                    pagination={false}
                    dataSource={[
                      {
                        key: 1,
                        need: "事件标题、正文、发布时间、来源",
                        source: "产业政策与新闻资讯库",
                        level: "内容级",
                        coverage: "完整覆盖",
                        auth: "当前场景可用",
                        action: "直接复用",
                      },
                      {
                        key: 2,
                        need: "项目名称、金额、中标主体、建设周期",
                        source: "重大项目库＋全国招投标公告库",
                        level: "事件＋主体",
                        coverage: "映射后覆盖",
                        auth: "当前场景可用",
                        action: "补充映射",
                      },
                      {
                        key: 3,
                        need: "企业主体ID、存量客户、管户关系",
                        source: "工商资源＋行内客户主数据",
                        level: "主体级",
                        coverage: "完整覆盖",
                        auth: "受控使用",
                        action: "直接复用",
                      },
                      {
                        key: 4,
                        need: "处罚正文、舆情摘要、风险传播范围",
                        source: "处罚库＋舆情资讯",
                        level: "字段＋授权",
                        coverage: "缺少正文摘要授权",
                        auth: "现有合同仅限查询",
                        action: "转授权扩围",
                      },
                    ]}
                    columns={[
                      { title: "需求字段", dataIndex: "need" },
                      { title: "候选已有资源", dataIndex: "source" },
                      { title: "比对粒度", dataIndex: "level" },
                      {
                        title: "覆盖结论",
                        dataIndex: "coverage",
                        render: (x: string) => (
                          <Tag color={x.includes("缺少") ? "red" : "green"}>
                            {x}
                          </Tag>
                        ),
                      },
                      { title: "授权核查", dataIndex: "auth" },
                      {
                        title: "处置方式",
                        dataIndex: "action",
                        render: (x: string) => <b className="red">{x}</b>,
                      },
                    ]}
                  />
                </Card>
                <div className="recommend-box">
                  <RobotOutlined />
                  <div>
                    <b>联合查重：规则比对＋语义比对＋样本剖析＋人工确认</b>
                    <p>
                      系统区分可直接复用、需字段映射和真实采购缺口；只有现有资源与授权无法满足的内容才形成采购事项。
                    </p>
                  </div>
                </div>
              </>
            ),
          },
        ]}
      />
      <div className="pane-actions">
        <Button onClick={() => go("procurement")}>查看关联采购事项</Button>
        <Button type="primary" onClick={next}>
          确认复用评估
        </Button>
      </div>
    </div>
  );
}

function PathChoice({ next }: { next: () => void }) {
  const [path, setPath] = useState(2);
  const items = [
    [
      "直接复用标准产品",
      "覆盖度 ≥ 95%",
      "1–2天",
      "95%–100%",
      "仅开通权限和配置机构",
    ],
    [
      "基于产品母版适配",
      "核心流程一致",
      "8–12天",
      "75%–90%",
      "参数调整、数据映射、本地配置",
    ],
    [
      "组合能力模块建设",
      "部分能力一致",
      "15–25天",
      "45%–75%",
      "重新编排模块与规则",
    ],
    ["新增开发", "现有能力不足", "30–60天", "0%–40%", "新建数据接入和处理逻辑"],
  ];
  return (
    <div className="build-pane">
      <Card size="small" className="period-note">
        <b>预计周期口径</b>
        <span>
          从建设路径确认后开始计算，包含配置开发、数据映射、联调测试和发布准备；不含外数采购、合同审批及新增资源接入周期。
        </span>
      </Card>
      <div className="path-grid">
        {items.map((x, i) => (
          <button
            key={x[0]}
            className={path === i ? "active" : ""}
            onClick={() => setPath(i)}
          >
            <span className="path-no">路径 {i + 1}</span>
            {path === i && <CheckCircleFilled />}
            <h3>{x[0]}</h3>
            <p>{x[1]}</p>
            <dl>
              <dt>预计建设周期</dt>
              <dd>{x[2]}</dd>
              <dt>预计复用比例</dt>
              <dd>{x[3]}</dd>
              <dt>主要工作</dt>
              <dd>{x[4]}</dd>
              <dt>维护责任</dt>
              <dd>{i === 0 ? "总行" : "总行母版＋属地配置"}</dd>
            </dl>
          </button>
        ))}
      </div>
      <Card className="scheme-summary">
        <SectionTitle title="总行标准化建设方案" sub="基于某分行半导体样本，复用母版与能力，依托全国企业库重新建设；不直接复制地方产品" />
        <div className="scheme-band">
          <div>
            <b>保留分行业务逻辑</b>
            <span>链属识别、机会信号、风险核验、企业名单</span>
          </div>
          <em>＋</em>
          <div>
            <b>全国化替代与参数化</b>
            <span>四川名录→全国企业库；地区、产业、环节和阈值转为参数</span>
          </div>
          <em>＋</em>
          <div>
            <b>总行新增建设</b>
            <span>事实链、产业链传导、综合研判、多岗位行动建议</span>
          </div>
        </div>
      </Card>
      <Card title="标准化前后建设效能对比" size="small">
        <Table pagination={false} rowKey="item" dataSource={[
          {item:"适用范围",before:"分行辖内半导体产业链",after:"全国地区 × 8类产业 × 多链属环节"},
          {item:"数据底座",before:"地方企业名录＋分行可用外数",after:"全国企业库＋全行公共外数＋行内客户数据"},
          {item:"整体复用率",before:"地方产品内部复用 42%",after:"母版与能力复用 78%"},
          {item:"开发周期",before:"同类产品逐省开发约30天",after:"标准产品建设18天；地方配置1—2天"},
          {item:"年度维护成本",before:"各分行分散维护，基准100",after:"总行统一维护，预计降至62"}
        ]} columns={[{title:"比较项",dataIndex:"item"},{title:"标准化前",dataIndex:"before"},{title:"标准化后",dataIndex:"after"}]} />
      </Card>
      <div className="pane-actions">
        <Button>保存选择</Button>
        <Button
          type="primary"
          onClick={() => {
            message.success("建设路径已确认，产品状态变更为“建设中”");
            next();
          }}
        >
          生成建设方案
        </Button>
      </div>
    </div>
  );
}

function Breakdown({ next }: { next: () => void }) {
  const rows = [
    [
      "业务任务",
      "产业事件融资机会与风险识别",
      "某分行半导体运行场景",
      "标准产品母版",
    ],
    [
      "数据输入",
      "政策＋项目＋招投标＋工商＋处罚＋司法＋舆情",
      "事件产品输入清单",
      "产品母版",
    ],
    ["加工处理", "事件抽取、去重与标准化", "事件标准化 M010", "能力模块"],
    ["加工处理", "企业与客户主体关联", "主体关联 M001", "能力模块"],
    ["加工处理", "机会/风险分类与影响评分", "事件判断规则组", "新增开发"],
    [
      "加工处理",
      "融资机会评分与风险交叉核验",
      "机会评分 M011＋司法核验 M005",
      "能力模块",
    ],
    ["结果交付", "事件卡片、机会清单、风险提示", "双结果输出组件", "产品母版"],
    ["后续动作", "营销与风险双任务分派、反馈回流", "流程分派 M012", "新增开发"],
  ].map((x, i) => ({
    key: i,
    section: x[0],
    node: x[1],
    source: x[2],
    type: x[3],
  }));
  return (
    <div className="build-pane">
      <Card>
        <Table
          pagination={false}
          dataSource={rows}
          columns={[
            {
              title: "五段链路",
              dataIndex: "section",
              render: (x: string) => <b>{x}</b>,
            },
            { title: "拆解节点", dataIndex: "node" },
            { title: "推荐来源", dataIndex: "source" },
            {
              title: "沉淀类型",
              dataIndex: "type",
              render: (x: string) => (
                <Select
                  style={{ width: 140 }}
                  defaultValue={x}
                  options={[
                    "产品母版",
                    "能力模块",
                    "通用规则",
                    "本地扩展",
                    "新增开发",
                  ].map((value) => ({ value }))}
                />
              ),
            },
            { title: "可复用", render: () => <Tag color="green">是</Tag> },
            {
              title: "维护主体",
              render: (_: any, r: any) =>
                r.type === "本地扩展" ? "浙江分行" : "总行数据管理部",
            },
          ]}
        />
      </Card>
      <div className="reuse-summary">
        <span>
          共拆解 <b>8</b> 个节点
        </span>
        <span>
          可直接复用 <b>5</b> 个
        </span>
        <span>
          母版扩展 <b>1</b> 个
        </span>
        <span>
          新增开发 <b>2</b> 个
        </span>
      </div>
      <div className="pane-actions">
        <Button>保存拆解</Button>
        <Button type="primary" onClick={next}>
          确认拆解并匹配能力
        </Button>
      </div>
    </div>
  );
}

function CapabilityMatch({ next }: { next: () => void }) {
  const [source, setSource] = useState<any>(null);
  const [selectedIds, setSelectedIds] = useState([
    "M010",
    "M001",
    "M011",
    "M005",
  ]);
  const recommendations = [
    {
      need: "将资讯转化为标准事件",
      asset: "事件抽取与标准化",
      id: "M010",
      score: 96,
      reason: "政策、项目与资讯事件结构可直接复用",
      selected: true,
    },
    {
      need: "企业和客户主体统一识别",
      asset: "主体关联",
      id: "M001",
      score: 98,
      reason: "输入输出完全一致",
      selected: true,
    },
    {
      need: "识别融资机会并形成优先级",
      asset: "融资机会评分",
      id: "M011",
      score: 91,
      reason: "复用既有产业事件产品评分能力",
      selected: true,
    },
    {
      need: "交叉核验司法与处罚风险",
      asset: "司法风险核验",
      id: "M005",
      score: 94,
      reason: "风险口径与全行标准一致",
      selected: true,
    },
    {
      need: "分别分派营销与风险任务",
      asset: "双任务分派与反馈",
      id: "M012",
      score: 62,
      reason: "现有任务推送仅支持单一业务流程，需新增",
      selected: true,
    },
  ];
  return (
    <div className="build-pane">
      <div className="match-head">
        <span>需求拆解结果</span>
        <SwapOutlined />
        <span>系统推荐能力资产</span>
      </div>
      {recommendations.map((x) => {
        const picked = selectedIds.includes(x.id);
        return (
          <Card key={x.id} className="match-row">
            <div className="need">
              <small>需求节点</small>
              <b>{x.need}</b>
            </div>
            <div className="match-link">
              <Progress percent={x.score} size="small" strokeColor="#c7000b" />
              <span>{x.reason}</span>
            </div>
            <div className="asset">
              <Tag color={x.score < 70 ? "orange" : "red"}>
                {x.score < 70 ? "待新增/改造" : "能力模块"}
              </Tag>
              <b>{x.asset}</b>
              <small>
                {x.id} · V2.3 · 已调用{" "}
                {modules.find((m) => m.id === x.id)?.calls || 0} 次
              </small>
            </div>
            <Space>
              <Button size="small" onClick={() => setSource(x)}>
                查看来源
              </Button>
              <Button
                size="small"
                type={picked ? "primary" : "default"}
                icon={picked ? <CheckCircleOutlined /> : <PlusOutlined />}
                onClick={() =>
                  setSelectedIds(
                    picked
                      ? selectedIds.filter((id) => id !== x.id)
                      : [...selectedIds, x.id],
                  )
                }
              >
                {picked ? "已选用" : "选用"}
              </Button>
            </Space>
          </Card>
        );
      })}
      <Card className="rule-match">
        <b>同步推荐通用规则</b>
        <Space wrap>
          {rules.slice(0, 8).map((r) => (
            <Tag
              key={r.id}
              closable
              color={
                ["R002", "R003", "R001", "R006"].includes(r.id)
                  ? "red"
                  : "default"
              }
            >
              {r.name}
            </Tag>
          ))}
        </Space>
      </Card>
      <div className="pane-actions">
        <Button
          onClick={() =>
            message.success(
              `匹配方案已保存：复用 ${selectedIds.length} 项，新增改造 ${recommendations.length - selectedIds.length} 项`,
            )
          }
        >
          保存匹配方案
        </Button>
        <Button type="primary" onClick={next}>
          进入流程编排
        </Button>
      </div>
      <Drawer
        width={620}
        open={!!source}
        onClose={() => setSource(null)}
        title="能力来源与复用边界"
      >
        {source && (
          <>
            <Descriptions
              bordered
              column={1}
              items={[
                {
                  label: "能力资产",
                  children: `${source.asset}（${source.id}）`,
                },
                {
                  label: "来源产品",
                  children:
                    source.id === "M011"
                      ? "产业事件驱动的融资机会识别 V2.1"
                      : "全行企业级公共能力库",
                },
                { label: "匹配需求", children: source.need },
                { label: "匹配依据", children: source.reason },
                {
                  label: "复用结论",
                  children:
                    source.score < 70
                      ? "需在既有能力基础上新增双任务路由与反馈契约"
                      : "标准输入输出一致，可直接引用",
                },
              ]}
            />
            <Card size="small" style={{ marginTop: 16 }} title="标准输入输出">
              <p>
                <b>输入：</b>标准事件、企业主体ID、客户归属与配置参数
              </p>
              <p>
                <b>输出：</b>评分/等级、判断依据、可解释标签与后续动作
              </p>
            </Card>
          </>
        )}
      </Drawer>
    </div>
  );
}

function SortableNode({
  node,
  onSelect,
  selected,
}: {
  node: any;
  onSelect: () => void;
  selected: boolean;
}) {
  const { attributes, listeners, setNodeRef, transform, transition } =
    useSortable({ id: node.id });
  return (
    <div
      ref={setNodeRef}
      style={{ transform: CSS.Transform.toString(transform), transition }}
      className={"flow-node " + (selected ? "selected" : "")}
      onClick={onSelect}
    >
      <span className="drag" {...attributes} {...listeners}>
        <DragOutlined />
      </span>
      <span className={"node-type " + node.type.replace(/\s/g, "")}>
        {node.type}
      </span>
      <span className="flow-node-content">
        <b>{node.name}</b>
        <small>{node.input || "企业主体ID"} → {node.output || `${node.name}结果`}</small>
        {!!node.rules?.length && <em>{node.rules.slice(0, 2).join(" · ")}{node.rules.length > 2 ? ` 等${node.rules.length}条` : ""}</em>}
      </span>
      <CheckCircleFilled />
    </div>
  );
}
function FlowCanvas({ next, profile }: { next: () => void; profile: any }) {
  const ruleOptions = ["地区范围规则", "行业准入规则", "注册年限规则", "机会评分阈值", "司法风险排除", "经营异常排除"];
  const buildFlow = () =>
    profile.flow.map((name: string, i: number) => ({
      id: `n${i + 1}`,
      name,
      type:
        i === 0
          ? "输入组件"
          : name.includes("评分") ||
              name.includes("分类") ||
              name.includes("核验")
            ? "决策能力"
            : name.includes("分派") || name.includes("输出")
              ? "输出组件"
              : i === 2
                ? "外数资源"
                : "能力模块",
      status: "ready",
      input: i === 0 ? "企业主体输入" : "企业主体ID",
      output: i === 0
        ? "企业主体输入结果"
        : name.includes("输出")
          ? "客户名单与行动建议"
          : `${name}结果`,
      rules: name.includes("分类") || name.includes("评分")
        ? ["地区范围规则", "行业准入规则", "机会评分阈值"]
        : name.includes("核验") ? ["司法风险排除", "经营异常排除"] : [],
      resources: name.includes("事件") ? ["EXT-R004"] : ["EXT-R001"],
      params: name.includes("评分") ? "机会分≥70；风险等级≤中" : "采用全行标准口径",
    }));
  const [flow, setFlow] = useState(buildFlow);
  const [selected, setSelected] = useState(buildFlow()[3]);
  const [draft, setDraft] = useState<any>({ ...buildFlow()[3] });
  const [checked, setChecked] = useState(false);
  const [agentOpen, setAgentOpen] = useState(false);
  const [drafted, setDrafted] = useState(false);
  const [agentPrompt, setAgentPrompt] = useState("请基于报送需求自动拆解流程，优先复用已有能力模块与通用规则，仅将地方差异保留为参数。");
  const [resourceSearch, setResourceSearch] = useState("");
  const [internalImported, setInternalImported] = useState(false);
  const [previewOpen, setPreviewOpen] = useState(false);
  const [previewRan, setPreviewRan] = useState(false);
  const [baselineCount, setBaselineCount] = useState(52);
  const [changeLog, setChangeLog] = useState<string[]>(["载入产业客户机会识别母版 V3.1"]);
  useEffect(() => {
    const fresh = buildFlow();
    setFlow(fresh);
    setSelected(fresh[3]);
    setDraft({ ...fresh[3] });
    setChecked(false);
  }, [profile.task]);
  const sensors = useSensors(useSensor(PointerSensor));
  const dragEnd = (e: DragEndEvent) => {
    const { active, over } = e;
    if (over && active.id !== over.id) {
      setFlow((items) =>
        arrayMove(
          items,
          items.findIndex((x) => x.id === active.id),
          items.findIndex((x) => x.id === over.id),
        ),
      );
    }
  };
  const add = (name: string, type: string) => {
    const node = { id: "n" + Date.now(), name, type, status: "ready", input: "企业主体ID", output: `${name}结果`, rules: [], resources: type === "外数资源" ? ["EXT-R001"] : [], params: "采用全行标准口径" };
    setFlow([...flow, node]);
    setSelected(node);
    setDraft({ ...node });
    message.success(`已添加节点：${name}`);
    setChangeLog((old) => [`增加${type}“${name}”`, ...old].slice(0, 5));
  };
  const applyTaskPack = () => {
    const packed = [
      { id: "tp1", name: "企业主体输入", type: "输入组件", status: "ready", input: "企业主体输入", output: "企业主体输入结果", rules: [], resources: [] },
      { id: "tp2", name: "主体关联与标准化", type: "能力模块", status: "ready", input: "企业主体输入结果", output: "标准企业主体ID", rules: [], resources: ["EXT-R001"], params: "统一社会信用代码优先；名称模糊匹配阈值≥90%" },
      { id: "tp3", name: "产业事件识别", type: "决策能力", status: "ready", input: "标准企业主体ID＋产业事件", output: "事件标签与影响方向", rules: ["行业准入规则"], resources: ["EXT-R004"], params: "政策、订单、扩产与价格事件；观察期近90日" },
      { id: "tp4", name: "机会与风险综合评分", type: "决策能力", status: "ready", input: "企业事实与事件标签", output: "机会指数＋风险指数＋判断依据", rules: ["地区范围规则", "行业准入规则", "机会评分阈值", "司法风险排除"], resources: ["EXT-R001", "EXT-R004"], params: "机会分≥70；风险等级≤中；保留证据来源" },
      { id: "tp5", name: "人工复核", type: "控制节点", status: "ready", input: "评分结果与证据链", output: "复核结论", rules: [], resources: [], params: "高风险与低置信度结果必须复核" },
      { id: "tp6", name: "名单、简报与任务分派", type: "输出组件", status: "ready", input: "复核结论", output: "企业名单＋行业简报＋营销/风险任务", rules: [], resources: [], params: "按客户经理、授信与风险岗位分别生成办公成果" },
    ];
    setFlow(packed);
    setSelected(packed[3]);
    setDraft({ ...packed[3], rules: [...packed[3].rules], resources: [...packed[3].resources] });
    setChecked(false);
    setChangeLog((old) => ["应用产业客户机会识别产品母版 V3.1", ...old].slice(0, 5));
    message.success("已应用“产业客户机会识别”产品母版：6个主流程节点，可继续增加或替换能力");
  };
  const selectNode = (node: any) => {
    setSelected(node);
    setDraft({ ...node, rules: [...(node.rules || [])], resources: [...(node.resources || [])] });
  };
  const saveNode = () => {
    const saved = { ...selected, ...draft };
    setFlow((items) => items.map((item) => item.id === saved.id ? saved : item));
    setSelected(saved);
    setDraft({ ...saved, rules: [...(saved.rules || [])], resources: [...(saved.resources || [])] });
    setChecked(false);
    message.success(`已保存“${saved.name}”，画布已同步更新`);
  };
  const replaceNode = (moduleName: string) => {
    const target = modules.find((x: any) => x.name === moduleName);
    if (!target) return;
    const replaced = { ...selected, name: target.name, type: "能力模块", input: target.input, output: target.output, params: "采用全行标准口径；保留原节点规则与异常处理", replacedFrom: selected.name };
    setFlow((items) => items.map((item) => item.id === selected.id ? replaced : item));
    setSelected(replaced);
    setDraft({ ...replaced, rules: [...(replaced.rules || [])], resources: [...(replaced.resources || [])] });
    setChecked(false);
    setPreviewRan(false);
    setChangeLog((old) => [`${selected.name} → ${target.name}`, ...old].slice(0, 5));
    message.success(`已用“${target.name}”替换“${selected.name}”，请重新校验并预览结果`);
  };
  const runPreview = () => {
    const delta = Math.max(3, Math.min(15, flow.length - 7 + new Set(flow.flatMap((x) => x.rules || [])).size));
    setBaselineCount(52 + delta);
    setPreviewRan(true);
    setPreviewOpen(true);
    message.success("已用同一批脱敏样本完成变更后预览");
  };
  return (
    <div className="build-pane canvas-layout">
      <Card className="resource-library" title="建设资源库" size="small">
        <Tabs
          size="small"
          items={[
            {
              key: "modules",
              label: "能力模块",
              children: (
                <List
                  size="small"
                  dataSource={modules}
                  renderItem={(x) => (
                    <List.Item
                      actions={[
                        <Button
                          key="a"
                          size="small"
                          type="text"
                          icon={<PlusOutlined />}
                          onClick={() => add(String(x.name), "能力模块")}
                        />,
                      ]}
                    >
                      <span>
                        <DeploymentUnitOutlined /> {x.name}
                        <small>
                          {x.input} → {x.output}
                        </small>
                      </span>
                    </List.Item>
                  )}
                />
              ),
            },
            {
              key: "resources",
              label: "外数资源",
              children: (
                <>
                  <Input
                    size="small"
                    allowClear
                    prefix={<SearchOutlined />}
                    value={resourceSearch}
                    onChange={(e) => setResourceSearch(e.target.value)}
                    placeholder="检索全部标准资源"
                  />
                  <List
                    size="small"
                    className="canvas-resource-list"
                    dataSource={resources.filter((x: any) =>
                      `${x.name}${x.id}${x.supplier}${x.fields}`
                        .toLowerCase()
                        .includes(resourceSearch.toLowerCase()),
                    )}
                    renderItem={(x) => (
                      <List.Item
                        actions={[
                          <Button
                            key="a"
                            size="small"
                            type="text"
                            icon={<PlusOutlined />}
                            onClick={() => add(x.name, "外数资源")}
                          />,
                        ]}
                      >
                        <span>
                          <DatabaseOutlined /> {x.name}
                          <small>
                            {x.id} · {x.supplier}
                          </small>
                        </span>
                      </List.Item>
                    )}
                  />
                </>
              ),
            },
            {
              key: "rules",
              label: "通用规则",
              children: (
                <List
                  size="small"
                  dataSource={rules}
                  renderItem={(x) => (
                    <List.Item
                      actions={[
                        <Button
                          key="a"
                          size="small"
                          type="text"
                          icon={<PlusOutlined />}
                          onClick={() => add(x.name, "通用规则")}
                        />,
                      ]}
                    >
                      <span>
                        <NodeIndexOutlined /> {x.name}
                        <small>
                          {x.input} → {x.output}
                        </small>
                      </span>
                    </List.Item>
                  )}
                />
              ),
            },
            {
              key: "internal",
              label: "导入内数",
              children: (
                <div className="internal-import">
                  <div className="internal-data-note">
                    仅在产品建设阶段导入已获授权的行内数据文件，完成字段识别与标准映射后供流程调用。
                  </div>
                  <Upload
                    beforeUpload={() => false}
                    showUploadList={false}
                    onChange={() => {
                      if (!internalImported)
                        add("行内客户经营数据", "内数输入");
                      setInternalImported(true);
                      message.success(
                        "已导入演示文件：识别6类行内字段，等待字段映射",
                      );
                    }}
                  >
                    <Button block icon={<CloudUploadOutlined />}>
                      {internalImported
                        ? "重新导入内数文件"
                        : "导入 Excel / CSV"}
                    </Button>
                  </Upload>
                  {internalImported && (
                    <div className="internal-import-result">
                      <Tag color="green">导入成功</Tag>
                      <b>客户经营融合样本.xlsx</b>
                      <small>
                        10,286行 ·
                        客户ID、管户机构、授信敞口、结算变化、产品持有、历史任务
                      </small>
                      <Button
                        size="small"
                        onClick={() =>
                          message.info(
                            "已生成6项标准字段映射，进入下一步可继续校验",
                          )
                        }
                      >
                        查看字段映射
                      </Button>
                    </div>
                  )}
                </div>
              ),
            },
            {
              key: "templates",
              label: "产品母版",
              children: (
                <div className="task-pack-picker">
                  <Tag color="red">推荐匹配 94%</Tag>
                  <h4>产业客户机会识别母版 V3.1</h4>
                  <p>复用成熟业务骨架、输入输出和组件关系；建设人员仍可逐节点增补、替换与配置。</p>
                  <div>
                    <span>6个主节点</span><span>4条规则</span><span>3类成果</span>
                  </div>
                  <Button block type="primary" icon={<RobotOutlined />} onClick={applyTaskPack}>
                    一键应用到画布
                  </Button>
                  <small>母版只提供骨架；地方扩展和新增能力需单独验证。</small>
                </div>
              ),
            },
            {
              key: "control",
              label: "输出组件",
              children: (
                <List
                  size="small"
                  dataSource={[
                    "客户名单输出",
                    "风险预警输出",
                    "准入结论输出",
                    "趋势预测输出",
                    "决策报告生成",
                    "策略建议输出",
                    "任务分派",
                    "人工复核",
                  ]}
                  renderItem={(x) => (
                    <List.Item
                      actions={[
                        <Button
                          key="a"
                          size="small"
                          type="text"
                          icon={<PlusOutlined />}
                          onClick={() =>
                            add(x, x.includes("复核") ? "控制节点" : "输出组件")
                          }
                        />,
                      ]}
                    >
                      {x}
                    </List.Item>
                  )}
                />
              ),
            },
          ]}
        />
        <Button block icon={<PlusOutlined />} onClick={() => add("未命名处理节点", "能力模块")}>
          新增空白节点
        </Button>
      </Card>
      <Card
        className="flow-canvas"
        title={
          <span>
            产品流程画布 <Tag>拖拽调整顺序</Tag>
          </span>
        }
        size="small"
        extra={
          <Space>
            <Button
              size="small"
              icon={<RobotOutlined />}
              onClick={() => setAgentOpen(true)}
            >
              智能编排助手
            </Button>
            <Button size="small" onClick={() => setFlow(buildFlow())}>
              自动布局
            </Button>
            <Button
              size="small"
              type="primary"
              onClick={() => {
                setChecked(true);
                message.success(
                  `运行检查通过：${flow.length}个节点、${flow.length - 1}条连接、无断点`,
                );
              }}
            >
              运行检查
            </Button>
            <Button size="small" onClick={runPreview}>预览结果</Button>
          </Space>
        }
      >
        <div className="flow-title-node">
          <b>业务需求</b>
          <span>{profile.task}</span>
        </div>
        <div className="vertical-line" />
        <DndContext
          sensors={sensors}
          collisionDetection={closestCenter}
          onDragEnd={dragEnd}
        >
          <SortableContext
            items={flow.map((x) => x.id)}
            strategy={verticalListSortingStrategy}
          >
            {flow.map((n, i) => (
              <React.Fragment key={n.id}>
                <SortableNode
                  node={n}
                  selected={selected.id === n.id}
                  onSelect={() => selectNode(n)}
                />
                {i < flow.length - 1 && <div className="node-arrow">↓</div>}
              </React.Fragment>
            ))}
          </SortableContext>
        </DndContext>
        {checked && (
          <div className="check-pass">
            <CheckCircleFilled /> 流程校验通过 · 输入输出完整 · 节点版本有效
          </div>
        )}
      </Card>
      <Card className="node-properties" title={<span>节点属性配置 <Tag color="red">保存后回写画布</Tag></span>} size="small">
        <Form layout="vertical" key={selected.id}>
          <div className="property-section-title">① 基础属性</div>
          <Form.Item label="节点名称">
            <Input value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} />
          </Form.Item>
          <Form.Item label="节点类型">
            <Select
              value={draft.type}
              onChange={(type) => setDraft({ ...draft, type })}
              options={[
                "能力模块",
                "决策能力",
                "外数资源",
                "内数输入",
                "通用规则",
                "控制节点",
                "输入组件",
                "输出组件",
              ].map((value) => ({ value }))}
            />
          </Form.Item>
          <Row gutter={8}>
            <Col span={12}>
              <Form.Item label="输入对象">
                <Input value={draft.input} onChange={(e) => setDraft({ ...draft, input: e.target.value })} />
              </Form.Item>
            </Col>
            <Col span={12}>
              <Form.Item label="输出对象">
                <Input value={draft.output} onChange={(e) => setDraft({ ...draft, output: e.target.value })} />
              </Form.Item>
            </Col>
          </Row>
          <div className="property-section-title">② 复用资产</div>
          <Form.Item label="替换为同契约能力" extra="替换后保留节点位置、规则和异常处理，并重新运行契约校验">
            <Select showSearch placeholder="选择替代模块" value={undefined} onChange={replaceNode} options={modules.slice(0, 18).map((x:any) => ({ value: x.name, label: `${x.name}（${x.id}）` }))} />
          </Form.Item>
          <Form.Item label="调用标准资源">
            <Select
              mode="multiple"
              value={draft.resources || []}
              onChange={(resources) => setDraft({ ...draft, resources })}
              options={resources
                .slice(0, 5)
                .map((x) => ({ value: x.id, label: x.name }))}
            />
          </Form.Item>
          <Form.Item label="挂载通用规则" extra="来自全行规则库，可跨产品复用并参数化调整">
            <Select mode="multiple" value={draft.rules || []} onChange={(rules) => setDraft({ ...draft, rules })}
              options={ruleOptions.map((value) => ({ value }))} placeholder="选择本节点调用的通用规则" />
          </Form.Item>
          <div className="property-section-title">③ 参数与输出</div>
          <Form.Item label="本节点参数口径">
            <Input.TextArea
              rows={3}
              value={draft.params}
              onChange={(e) => setDraft({ ...draft, params: e.target.value })}
              placeholder="例如：近90日；机会分≥70；排除近3个月重大司法风险"
            />
          </Form.Item>
          <Form.Item label="结果表达">
            <Select
              defaultValue="结论＋评分＋判断依据＋建议动作"
              options={[
                { value: "结论＋评分＋判断依据＋建议动作" },
                { value: "名单＋排序＋推荐理由" },
                { value: "风险等级＋命中规则＋处置建议" },
              ]}
            />
          </Form.Item>
          <Form.Item label="失败处理方式">
            <Select
              defaultValue="记录异常并进入人工核验"
              options={[
                { value: "记录异常并进入人工核验" },
                { value: "终止流程" },
              ]}
            />
          </Form.Item>
          <Space className="node-property-actions">
            <Button
              type="primary"
              onClick={saveNode}
            >
              保存并更新画布
            </Button>
            <Button
              danger
              onClick={() => setFlow(flow.filter((x) => x.id !== selected.id))}
            >
              删除
            </Button>
          </Space>
        </Form>
      </Card>
      <div className="canvas-footer">
        <Space>
          <span>节点 {flow.length}</span>
          <span>连接 {Math.max(0, flow.length - 1)}</span>
          <span>能力模块 {flow.filter((x) => x.type === "能力模块" || x.type === "决策能力").length}</span>
          <span>通用规则 {new Set(flow.flatMap((x) => x.rules || [])).size}</span>
          <span>预计复用率 68%</span>
        </Space>
        <Space>
          <Button onClick={() => message.success("流程方案已保存")}>
            保存方案
          </Button>
          <Button
            onClick={() =>
              message.success(
                "组装清单已生成：资源、能力、规则、决策输出与责任边界已列明",
              )
            }
          >
            生成组装清单
          </Button>
          <Button onClick={runPreview}>运行变更预览</Button>
          <Button type="primary" disabled={!checked} onClick={next}>
            完成编排，进入映射
          </Button>
        </Space>
      </div>
      <Modal
        width={1160}
        open={previewOpen}
        onCancel={() => setPreviewOpen(false)}
        title="产品变更影响与结果预览"
        footer={<Space><Button onClick={() => setPreviewOpen(false)}>返回画布继续调整</Button><Button type="primary" onClick={() => { setPreviewOpen(false); setChecked(true); message.success("预览结果已确认，当前组装版本可进入映射与验证"); }}>确认本次组装结果</Button></Space>}
      >
        <div className="preview-diff-head"><div><Tag>基线版本</Tag><b>母版 V3.1</b><small>8个节点 · 52家候选企业</small></div><em>→</em><div><Tag color="red">当前草案</Tag><b>ASSEMBLY-DRAFT</b><small>{flow.length}个节点 · {previewRan ? baselineCount : "待运行"}家候选企业</small></div><span><b>{changeLog.length}</b><small>本轮变更</small></span></div>
        <Row gutter={14}>
          <Col span={15}><Card size="small" title="同一批样本运行结果对比"><div className="preview-metrics"><span><small>候选企业</small><b>52 → {baselineCount}</b></span><span><small>机会信号</small><b>31 → {Math.max(34, baselineCount - 17)}</b></span><span><small>风险排除</small><b>9 → {Math.max(10, flow.filter(x => x.name.includes("风险") || x.name.includes("司法")).length + 9)}</b></span><span><small>待人工核验</small><b>12 → 8</b></span></div><Table size="small" pagination={false} rowKey="name" dataSource={companyResults.slice(0,4)} columns={[{title:"企业",dataIndex:"name"},{title:"基线结果",render:()=>"观察"},{title:"当前结果",render:(_:any,r:any)=><Tag color={r.score>80?"green":"blue"}>{r.score>80?"优先跟进":"纳入名单"}</Tag>},{title:"变化原因",render:()=>changeLog[0] || "参数更新"},{title:"证据链",render:()=>"资源＋模块＋规则可追溯"}]} /></Card></Col>
          <Col span={9}><Card size="small" title="变更影响清单"><div className="change-impact-list">{changeLog.map((x,i)=><span key={`${x}-${i}`}><CheckCircleFilled /><b>{x}</b><small>{i===0?"影响当前节点及下游名单、评分和行动提示":"已记录版本差异，可回滚"}</small></span>)}</div><div className="preview-gates"><b>进入下一阶段前</b><span>✓ 输入输出契约重新校验</span><span>✓ 代表性样本结果已对比</span><span>✓ 规则和资源版本已锁定</span><span>✓ 新增能力标记为回沉候选</span></div></Card></Col>
        </Row>
      </Modal>
      <Modal
        width={760}
        open={agentOpen}
        onCancel={() => setAgentOpen(false)}
        title={
          <Space>
            <RobotOutlined />
            智能编排助手 <DemoTag />
          </Space>
        }
        footer={
          <Space>
            <Button onClick={() => setAgentOpen(false)}>取消</Button>
            <Button
              type="primary"
              onClick={() => {
                setDrafted(true);
                const generated = buildFlow().map((node: any, index: number) => ({
                  ...node,
                  rules:
                    node.rules?.length
                      ? node.rules
                      : index === 0
                        ? ["地区范围规则", "行业准入规则"]
                        : node.name.includes("筛选")
                          ? ["地区范围规则", "行业准入规则", "机会评分阈值"]
                          : [],
                  params:
                    index === 0
                      ? "企业主体输入；地域和行业作为运行参数"
                      : node.params,
                  aiGenerated: true,
                }));
                setFlow(generated);
                setSelected(generated[0]);
                setDraft({ ...generated[0] });
                setChecked(false);
                message.success("编排草案已应用到画布，请确认节点与规则");
              }}
            >
              生成并应用编排草案
            </Button>
          </Space>
        }
      >
        <div className="agent-chat">
          <div className="ai-orchestration-labels">
            <Tag color="red">系统生成草案</Tag>
            <span>自然语言需求 → 节点拆解 → 资产匹配 → 规则挂载 → 参数建议</span>
            <Tag>人工确认后生效</Tag>
          </div>
          <Input.TextArea
            rows={3}
            value={agentPrompt}
            onChange={(e) => setAgentPrompt(e.target.value)}
            placeholder="例如：优先复用已有模块，挂载地区、行业和风险排除规则，输出企业名单与行动建议"
          />
          <p>
            <b>报送需求：</b>
            {profile.task}。
          </p>
          <div>
            <b>助手识别结果</b>
            <Row gutter={[10, 10]}>
              {[
                ["业务分类", profile.category],
                ["推荐资源", profile.inputs],
                ["推荐加工", profile.processing],
                ["交付结果", profile.output],
              ].map((x) => (
                <Col span={12} key={x[0]}>
                  <Card size="small">
                    <small>{x[0]}</small>
                    <b>{x[1]}</b>
                  </Card>
                </Col>
              ))}
            </Row>
            <p className="muted">
              系统依据当前候选产品报送信息生成编排草案；建设人员仍需确认资源授权、节点版本、参数口径和责任归属。
            </p>
            {drafted && <StatusTag status="草案已应用至画布，待人工确认" />}
          </div>
        </div>
      </Modal>
    </div>
  );
}

function DataMapping({ next }: { next: () => void }) {
  const [validated, setValidated] = useState(false);
  const eventRows = [
    {
      source: "政策资讯.标题/正文",
      standard: "事件标题/事件摘要",
      type: "文本",
      rule: "事件抽取与标准化 M010",
      key: "否",
      status: "待校验",
    },
    {
      source: "项目公告.项目名称",
      standard: "标准项目名称",
      type: "字符串",
      rule: "项目名称清洗规则",
      key: "否",
      status: "待校验",
    },
    {
      source: "招中标公告.中标单位",
      standard: "企业主体ID",
      type: "字符串",
      rule: "主体关联 M001",
      key: "是",
      status: "待校验",
    },
    {
      source: "处罚公告.发布日期",
      standard: "事件发生/披露时间",
      type: "日期",
      rule: "时间口径转换",
      key: "否",
      status: "待校验",
    },
    {
      source: "行内客户.客户编号",
      standard: "客户主体ID/管户机构",
      type: "字符串",
      rule: "客户主数据映射",
      key: "是",
      status: "待校验",
    },
  ];
  return (
    <div className="build-pane">
      <Card>
        <SectionTitle
          title="字段与主体映射"
          sub="将供应商、地方文件和接口字段统一映射至企业级标准字段"
          extra={
            <Space>
              <Button icon={<ReloadOutlined />}>自动匹配</Button>
              <Button
                type="primary"
                onClick={() => {
                  setValidated(true);
                  message.success("映射校验完成：5项通过，0项阻断");
                }}
              >
                校验映射
              </Button>
            </Space>
          }
        />
        <Table
          pagination={false}
          dataSource={eventRows}
          rowKey="source"
          columns={[
            {
              title: "来源字段",
              dataIndex: "source",
              render: (x: string) => <code>{x}</code>,
            },
            {
              title: "标准字段",
              dataIndex: "standard",
              render: (x: string) => (
                <Select
                  defaultValue={x}
                  style={{ width: 160 }}
                  options={[
                    { value: x },
                    { value: "标准企业名称" },
                    { value: "企业主体ID" },
                  ]}
                />
              ),
            },
            { title: "数据类型", dataIndex: "type" },
            { title: "转换规则", dataIndex: "rule" },
            { title: "主键", dataIndex: "key" },
            {
              title: "映射状态",
              dataIndex: "status",
              render: (x: string) => (
                <StatusTag status={validated ? "已匹配" : x} />
              ),
            },
            {
              title: "操作",
              render: () => <Button type="link">查看样例</Button>,
            },
          ]}
        />
      </Card>
      <Card className="lineage-strip">
        <b>映射进入产品流程</b>
        <span>多源公开资讯</span>
        <em>→</em>
        <span>标准事件ID</span>
        <em>→</em>
        <span>企业/客户主体ID</span>
        <em>→</em>
        <span>机会与风险信号</span>
      </Card>
      <div className="pane-actions">
        <Button onClick={() => message.success("事件映射版本 V1.0 已保存")}>
          保存版本
        </Button>
        <Button type="primary" disabled={!validated} onClick={next}>
          确认映射并配置规则
        </Button>
      </div>
    </div>
  );
}

function RuleConfig({ next }: { next: () => void }) {
  const [conds, setConds] = useState([
    ["覆盖地域", "属于", "四川省"],
    ["事件时效", "在近", "90日"],
    ["来源可信度", "不低于", "权威公开来源"],
    ["机会事件类型", "包含", "项目中标/扩产技改/政策支持"],
    ["机会评分", "不低于", "70分"],
    ["风险事件等级", "达到", "中风险及以上"],
    ["客户匹配", "优先", "存量及潜在客户"],
  ]);
  return (
    <div className="build-pane">
      <Row gutter={16}>
        <Col span={16}>
          <Card
            title="可视化条件配置器"
            extra={
              <Button
                size="small"
                onClick={() => setConds([...conds, ["经营异常", "等于", "无"]])}
                icon={<PlusOutlined />}
              >
                新增条件
              </Button>
            }
          >
            <div className="rule-builder">
              <div className="logic-main">
                满足以下{" "}
                <Select
                  size="small"
                  defaultValue="全部（且）"
                  options={[{ value: "全部（且）" }, { value: "任一（或）" }]}
                />{" "}
                条件
              </div>
              {conds.map((c, i) => (
                <div className="condition" key={i}>
                  <span>{i === 0 ? "当" : "且"}</span>
                  <Select value={c[0]} options={[{ value: c[0] }]} />
                  <Select value={c[1]} options={[{ value: c[1] }]} />
                  <Input
                    value={c[2]}
                    onChange={(e) =>
                      setConds(
                        conds.map((x, j) =>
                          j === i ? [x[0], x[1], e.target.value] : x,
                        ),
                      )
                    }
                  />
                  <Button
                    type="text"
                    danger
                    onClick={() => setConds(conds.filter((_, j) => j !== i))}
                  >
                    删除
                  </Button>
                </div>
              ))}
            </div>
          </Card>
        </Col>
        <Col span={8}>
          <Card title="已调用规则">
            <List
              dataSource={rules.slice(0, 7)}
              renderItem={(r) => (
                <List.Item
                  actions={[
                    <Button key="c" type="link" size="small">
                      替换
                    </Button>,
                  ]}
                >
                  <List.Item.Meta
                    avatar={<CheckCircleFilled className="green" />}
                    title={r.name}
                    description={`${r.id} · ${r.version}`}
                  />
                </List.Item>
              )}
            />
            <Button block icon={<PlusOutlined />}>
              调用已有规则
            </Button>
          </Card>
        </Col>
      </Row>
      <Card className="assembly">
        <SectionTitle title="产品组装清单" sub="根据流程、映射和规则自动生成" />
        <Row gutter={12}>
          {[
            ["外数资源", "7项", "直接复用"],
            ["既有产品", "1项", "能力复用"],
            ["能力模块", "5项", "直接复用"],
            ["通用规则", "6项", "参数适配"],
            ["流程配置", "2项", "重新编排"],
            ["新增开发", "2项", "新增开发"],
          ].map((x) => (
            <Col span={4} key={x[0]}>
              <div className="assembly-item">
                <small>{x[0]}</small>
                <b>{x[1]}</b>
                <Tag
                  color={
                    x[2] === "直接复用"
                      ? "green"
                      : x[2] === "新增开发"
                        ? "red"
                        : "orange"
                  }
                >
                  {x[2]}
                </Tag>
              </div>
            </Col>
          ))}
        </Row>
        <div className="assembly-metrics">
          <Progress
            type="dashboard"
            percent={68}
            size={110}
            strokeColor="#c7000b"
          />
          <div>
            <b>整体复用率 68%</b>
            <span>模块复用率 80% · 规则复用率 71%</span>
          </div>
          <div>
            <b>预计建设周期 22个工作日</b>
            <span>较完全新增开发缩短约 28 个工作日</span>
          </div>
          <div>
            <b>新增开发聚焦 2 项</b>
            <span>事件影响评分与双任务分派反馈</span>
          </div>
        </div>
      </Card>
      <div className="pane-actions">
        <Button onClick={() => message.success("事件判断规则组已保存")}>
          保存规则组
        </Button>
        <Button type="primary" onClick={next}>
          确认规则并进入验证
        </Button>
      </div>
    </div>
  );
}

const productTemplates = [
  {
    name: "先进制造产业链协同经营母版",
    scene: "设备更新、重大订单、原材料波动",
    core: "事件解析＋产业链传导＋融资匹配",
    status: "在建主母版",
  },
  {
    name: "农业经营主体增信母版",
    scene: "种养殖、农资、县域特色产业",
    core: "产能核验＋价格气象＋现金流估算",
    status: "可适配",
  },
  {
    name: "物流运力经营评价母版",
    scene: "公路货运、港航、园区物流",
    core: "轨迹运单＋运力指数＋动态风险",
    status: "可适配",
  },
  {
    name: "科技企业成长评价母版",
    scene: "专精特新、科创园区、研发企业",
    core: "专利人才＋订单融资＋成长阶段",
    status: "稳定运行",
  },
  {
    name: "绿色转型融资识别母版",
    scene: "节能改造、清洁能源、碳减排",
    core: "项目识别＋减排测算＋绿色认定",
    status: "待验证",
  },
  {
    name: "跨境经营与风险监测母版",
    scene: "跨境贸易、出海建厂、汇率波动",
    core: "海关物流＋国别风险＋结算机会",
    status: "待验证",
  },
];

function IndustryModel({ next }: { next: () => void }) {
  const [template, setTemplate] = useState(productTemplates[0].name);
  const [version, setVersion] = useState("V3.2 候选版");
  const selected = productTemplates.find((x) => x.name === template)!;
  const factors = [
    {
      name: "事件影响强度",
      weight: 24,
      evidence: "政策层级、项目金额、订单可信度",
    },
    {
      name: "产业链传导度",
      weight: 22,
      evidence: "供需关系、替代性、价格传导期限",
    },
    {
      name: "企业承接能力",
      weight: 21,
      evidence: "产能、专利、招聘、招投标与物流",
    },
    {
      name: "我行业务适配",
      weight: 18,
      evidence: "管户、授信、结算与产品覆盖",
    },
    { name: "风险约束", weight: 15, evidence: "司法、舆情、成本与回款异常" },
  ];
  return (
    <div className="build-pane standard-model">
      <SectionTitle
        title="行业模型与产品母版"
        sub="总行维护行业知识、特征口径和决策输出；分行仅配置地域、产业参数和授权数据"
        extra={
          <Space>
            <Tag color="red">总行标准产品</Tag>
            <Select
              value={version}
              onChange={setVersion}
              style={{ width: 150 }}
              options={["V3.2 候选版", "V3.1 生产版", "V3.0 历史版"].map(
                (value) => ({ value }),
              )}
            />
          </Space>
        }
      />
      <Row gutter={14}>
        <Col span={8}>
          <Card title="母版资产库" className="template-list">
            <List
              dataSource={productTemplates}
              renderItem={(x) => (
                <List.Item
                  className={template === x.name ? "active" : ""}
                  onClick={() => setTemplate(x.name)}
                >
                  <List.Item.Meta title={x.name} description={x.scene} />
                  <StatusTag status={x.status} />
                </List.Item>
              )}
            />
          </Card>
        </Col>
        <Col span={16}>
          <Card
            title={selected.name}
            extra={<Tag color="green">行业专家已复核</Tag>}
          >
            <Descriptions
              size="small"
              column={2}
              items={[
                { label: "适用场景", children: selected.scene },
                { label: "核心加工", children: selected.core },
                { label: "模型责任人", children: "总行公司金融部＋数据管理部" },
                { label: "更新机制", children: "事件触发重评＋季度回溯" },
              ]}
            />
            <h3>决策因子与证据口径</h3>
            <Table
              size="small"
              pagination={false}
              rowKey="name"
              dataSource={factors}
              columns={[
                {
                  title: "决策因子",
                  dataIndex: "name",
                  render: (x: string) => <b>{x}</b>,
                },
                {
                  title: "权重",
                  dataIndex: "weight",
                  render: (x: number) => (
                    <Progress percent={x} size="small" format={() => `${x}%`} />
                  ),
                },
                { title: "可核验证据", dataIndex: "evidence" },
                {
                  title: "治理状态",
                  render: () => <Tag color="green">口径已锁定</Tag>,
                },
              ]}
            />
            <div className="model-boundary">
              <b>标准化边界</b>
              <span>可配置：区域、产业目录、观察期限、阈值和任务路由</span>
              <span>
                总行锁定：主体口径、核心特征、证据等级、风险底线和模型解释
              </span>
            </div>
          </Card>
        </Col>
      </Row>
      <div className="pane-actions">
        <Button>保存模型草稿</Button>
        <Button type="primary" onClick={next}>
          进入策略仿真
        </Button>
      </div>
    </div>
  );
}

function StrategySimulation({ next }: { next: () => void }) {
  const [strategy, setStrategy] = useState("平衡经营策略");
  const [ran, setRan] = useState(false);
  const scenarios: any = {
    稳健风险策略: ["186家", "42家", "8.1%", "2.6%", "+19天"],
    平衡经营策略: ["326家", "79家", "13.8%", "4.7%", "+15天"],
    积极拓展策略: ["512家", "126家", "17.2%", "8.9%", "+11天"],
  };
  const vals = scenarios[strategy];
  return (
    <div className="build-pane">
      <SectionTitle
        title="策略仿真与业务门槛"
        sub="发布前使用历史样本和压力事件检验线索产出、风险识别与业务承载能力"
      />
      <Card className="simulation-bar">
        <Space wrap>
          <Select
            value={strategy}
            onChange={setStrategy}
            style={{ width: 190 }}
            options={Object.keys(scenarios).map((value) => ({ value }))}
          />
          <Select
            defaultValue="先进制造企业 12,680家"
            style={{ width: 220 }}
            options={[
              { value: "先进制造企业 12,680家" },
              { value: "存量授信客户 3,240家" },
            ]}
          />
          <Select
            defaultValue="回溯窗口：近12个月"
            style={{ width: 190 }}
            options={[
              { value: "回溯窗口：近12个月" },
              { value: "回溯窗口：近24个月" },
            ]}
          />
          <Button
            type="primary"
            icon={<ExperimentOutlined />}
            onClick={() => {
              setRan(true);
              message.success("策略仿真完成，已生成对比结果");
            }}
          >
            运行仿真
          </Button>
        </Space>
      </Card>
      <Row gutter={[12, 12]}>
        {[
          ["机会线索", vals[0]],
          ["建议触达", vals[1]],
          ["历史转化率", vals[2]],
          ["风险误报率", vals[3]],
          ["预警提前量", vals[4]],
        ].map((x) => (
          <Col span={Math.floor(24 / 5)} key={x[0]}>
            <MetricCard
              title={x[0]}
              value={x[1]}
              trend={ran ? "本轮仿真结果" : "历史基准测算"}
            />
          </Col>
        ))}
      </Row>
      <Row gutter={14}>
        <Col span={14}>
          <Card title="三类策略效果比较">
            <ReactECharts
              style={{ height: 300 }}
              option={{
                tooltip: { trigger: "axis" },
                legend: { data: ["线索转化率", "误报率", "任务完成率"] },
                xAxis: { type: "category", data: Object.keys(scenarios) },
                yAxis: { type: "value", axisLabel: { formatter: "{value}%" } },
                series: [
                  {
                    name: "线索转化率",
                    type: "bar",
                    data: [9.2, 13.8, 17.2],
                    itemStyle: { color: "#c7000b" },
                  },
                  {
                    name: "误报率",
                    type: "line",
                    data: [2.6, 4.7, 8.9],
                    itemStyle: { color: "#d98700" },
                  },
                  {
                    name: "任务完成率",
                    type: "line",
                    data: [91, 87, 72],
                    itemStyle: { color: "#237a57" },
                  },
                ],
              }}
            />
          </Card>
        </Col>
        <Col span={10}>
          <Card title="发布门禁检查">
            <List
              dataSource={[
                "证据可追溯率 ≥ 95%",
                "高风险样本召回率 ≥ 85%",
                "误报率 ≤ 6%",
                "任务日均增量不超过岗位承载上限",
                "模型解释与人工复核路径完整",
              ]}
              renderItem={(x, i) => (
                <List.Item>
                  <Space>
                    <CheckCircleFilled
                      className={
                        i === 3 && strategy === "积极拓展策略"
                          ? "warn"
                          : "positive"
                      }
                    />
                    <span>{x}</span>
                  </Space>
                  <Tag
                    color={
                      i === 3 && strategy === "积极拓展策略"
                        ? "orange"
                        : "green"
                    }
                  >
                    {i === 3 && strategy === "积极拓展策略" ? "需降载" : "通过"}
                  </Tag>
                </List.Item>
              )}
            />
          </Card>
        </Col>
      </Row>
      <div className="pane-actions">
        <Button>保存仿真报告</Button>
        <Button type="primary" onClick={next}>
          进入测试验证
        </Button>
      </div>
    </div>
  );
}

function ModelVersion({ next }: { next: () => void }) {
  return (
    <div className="build-pane">
      <SectionTitle
        title="模型与产品版本管理"
        sub="数据、特征、规则、模型、界面和嵌入接口共同组成一个可回滚的产品版本"
      />
      <Card>
        <div className="version-compare">
          <div>
            <Tag color="green">生产版本</Tag>
            <h2>V3.1</h2>
            <p>2026-07-15发布 · 覆盖先进制造与新能源装备</p>
            <ul>
              <li>事件模型 EM-2.4</li>
              <li>机会策略 OS-3.1</li>
              <li>风险策略 RS-2.8</li>
              <li>全行32家机构启用</li>
            </ul>
          </div>
          <div>
            <Tag color="red">待发布版本</Tag>
            <h2>V3.2</h2>
            <p>拟于2026-08-10发布 · 新增重大订单与成本传导</p>
            <ul>
              <li>事件模型 EM-2.6</li>
              <li>机会策略 OS-3.3</li>
              <li>风险策略 RS-3.0</li>
              <li>新增融e慧任务卡接口</li>
            </ul>
          </div>
        </div>
        <Timeline
          items={[
            { color: "green", children: "V3.2完成历史回溯与压力测试" },
            { color: "green", children: "模型、规则、字段映射已形成变更清单" },
            { color: "blue", children: "等待业务、风险与数据管理联合审批" },
            { color: "gray", children: "灰度发布后保留V3.1一键回滚能力" },
          ]}
        />
      </Card>
      <div className="pane-actions">
        <Button>查看版本差异</Button>
        <Button type="primary" onClick={next}>
          进入发布准备
        </Button>
      </div>
    </div>
  );
}

function EffectAttribution({ next }: { next: () => void }) {
  const metrics = [
    ["节省人工工时", "1,286小时", "较基线 +31%"],
    ["线索转化率", "13.8%", "较人工名单 +5.6pct"],
    ["授信转化率", "6.9%", "审批中口径"],
    ["预警提前量", "15天", "较规则监测 +8天"],
    ["误报率", "4.7%", "低于门限6%"],
  ];
  return (
    <div className="build-pane">
      <SectionTitle
        title="效果归因与迭代决策"
        sub="区分数据、模型、任务执行和外部环境的贡献，避免仅用访问量评价产品"
      />
      <div className="effect-metrics">
        {metrics.map((x) => (
          <Card key={x[0]}>
            <Statistic title={x[0]} value={x[1]} />
            <small>{x[2]}</small>
          </Card>
        ))}
      </div>
      <Row gutter={14}>
        <Col span={14}>
          <Card title="业务结果漏斗">
            <div className="effect-funnel">
              {[
                ["识别对象", "12,680"],
                ["有效线索", "326"],
                ["已分派", "286"],
                ["已触达", "214"],
                ["融资受理", "79"],
                ["授信通过", "42"],
              ].map((x, i) => (
                <React.Fragment key={x[0]}>
                  <div>
                    <span>{x[0]}</span>
                    <b>{x[1]}</b>
                  </div>
                  {i < 5 && <em>→</em>}
                </React.Fragment>
              ))}
            </div>
          </Card>
        </Col>
        <Col span={10}>
          <Card title="增量效果归因">
            <ReactECharts
              style={{ height: 250 }}
              option={{
                tooltip: { trigger: "item" },
                series: [
                  {
                    type: "pie",
                    radius: ["45%", "72%"],
                    data: [
                      { name: "外数新增证据", value: 34 },
                      { name: "行业模型优化", value: 27 },
                      { name: "任务路由优化", value: 21 },
                      { name: "业务执行差异", value: 12 },
                      { name: "外部环境", value: 6 },
                    ],
                    color: [
                      "#c7000b",
                      "#e04b52",
                      "#d98700",
                      "#4d6f8f",
                      "#aab2bd",
                    ],
                  },
                ],
              }}
            />
          </Card>
        </Col>
      </Row>
      <Card title="迭代建议" className="iteration-note">
        <b>建议将“重大订单真实性核验”沉淀为通用能力模块</b>
        <span>
          其对授信转化提升贡献显著，并已在四川、江苏、浙江三个样本中稳定运行；下一版本将降低原材料价格短期波动权重，减少风险误报。
        </span>
      </Card>
      <div className="pane-actions">
        <Button>生成运营复盘</Button>
        <Button type="primary" onClick={next}>
          完成归因并回沉能力
        </Button>
      </div>
    </div>
  );
}

function Testing({ next }: { next: () => void }) {
  const [running, setRunning] = useState(false);
  const [passed, setPassed] = useState(
    () =>
      typeof window !== "undefined" &&
      localStorage.getItem("icbc-build-tested") === "true",
  );
  const tests = [
    "政策、项目与资讯资源可正常获取",
    "事件抽取准确率与重复事件合并",
    "企业主体关联准确率",
    "融资机会评分规则符合预期",
    "风险事件交叉核验准确",
    "营销与风险双任务分派成功",
    "反馈结果可回流运营评价",
  ].map((x, i) => ({
    key: i,
    type: i < 5 ? "功能验证" : "运行验证",
    case: x,
    expect:
      i === 1
        ? "事件抽取F1 ≥ 90%"
        : i === 2
          ? "主体关联准确率 ≥ 98%"
          : "结果符合预期",
    actual: passed ? (i === 1 ? "92.6%" : i === 2 ? "99.1%" : "符合预期") : "—",
    status: passed ? "通过" : "待测试",
    owner: i >= 5 ? "四川分行业务验证组" : "系统自动",
  }));
  const run = () => {
    setRunning(true);
    setTimeout(() => {
      setRunning(false);
      setPassed(true);
      localStorage.setItem("icbc-build-tested", "true");
      message.success("6项测试全部通过，产品状态变更为“待发布”");
    }, 900);
  };
  const trace = [
    {
      node: "多源事件采集",
      api: "GET /mock-api/events",
      before: 28640,
      after: 28640,
      ms: 184,
    },
    {
      node: "事件抽取与去重",
      api: "ABILITY M010",
      before: 28640,
      after: 9142,
      ms: 426,
    },
    {
      node: "事件标准化",
      api: "POST /mock-api/event-normalize",
      before: 9142,
      after: 8876,
      ms: 218,
    },
    {
      node: "企业主体关联",
      api: "ABILITY M001",
      before: 8876,
      after: 6420,
      ms: 193,
    },
    {
      node: "机会与风险分类",
      api: "RULE EVENT-GROUP",
      before: 6420,
      after: 1248,
      ms: 86,
    },
    {
      node: "影响评分与客户匹配",
      api: "ABILITY M011",
      before: 1248,
      after: 373,
      ms: 142,
    },
    {
      node: "双任务分派",
      api: "ABILITY M012",
      before: 373,
      after: 373,
      ms: 64,
    },
  ];
  return (
    <div className="build-pane">
      <Card>
        <Tabs
          items={[
            {
              key: "1",
              label: "功能与适配验证",
              children: (
                <Table
                  pagination={false}
                  dataSource={tests}
                  columns={[
                    { title: "验证类型", dataIndex: "type" },
                    { title: "测试用例", dataIndex: "case" },
                    { title: "预期结果", dataIndex: "expect" },
                    { title: "实际结果", dataIndex: "actual" },
                    {
                      title: "状态",
                      dataIndex: "status",
                      render: (x: string) => <StatusTag status={x} />,
                    },
                    { title: "责任人", dataIndex: "owner" },
                  ]}
                />
              ),
            },
            {
              key: "2",
              label: "虚拟API试运行",
              children: !passed ? (
                <Empty description="点击下方“运行完整流程”，查看各API节点的数据收敛过程" />
              ) : (
                <>
                  <div className="run-funnel">
                    {trace.map((x, i) => (
                      <React.Fragment key={x.node}>
                        <div>
                          <small>{x.api}</small>
                          <b>{x.node}</b>
                          <strong>{x.after.toLocaleString()} 家</strong>
                          <span>{x.ms} ms · 成功</span>
                        </div>
                        {i < trace.length - 1 && <em>→</em>}
                      </React.Fragment>
                    ))}
                  </div>
                  <Table
                    size="small"
                    pagination={false}
                    dataSource={trace}
                    rowKey="node"
                    columns={[
                      { title: "运行节点", dataIndex: "node" },
                      {
                        title: "虚拟调用",
                        dataIndex: "api",
                        render: (x: string) => <code>{x}</code>,
                      },
                      {
                        title: "输入",
                        dataIndex: "before",
                        render: (x: number) => x + " 家",
                      },
                      {
                        title: "输出",
                        dataIndex: "after",
                        render: (x: number) => <b>{x} 家</b>,
                      },
                      {
                        title: "耗时",
                        dataIndex: "ms",
                        render: (x: number) => x + " ms",
                      },
                      {
                        title: "状态",
                        render: () => <StatusTag status="成功" />,
                      },
                    ]}
                  />
                </>
              ),
            },
            {
              key: "3",
              label: "结果预览",
              children: passed ? (
                <Table
                  size="small"
                  pagination={{ pageSize: 5 }}
                  dataSource={companyResults}
                  columns={[
                    { title: "企业名称", dataIndex: "name" },
                    { title: "地区", dataIndex: "city" },
                    { title: "资质", dataIndex: "qualification" },
                    { title: "招投标", dataIndex: "bids" },
                    { title: "推荐分", dataIndex: "score" },
                    { title: "风险", dataIndex: "risk" },
                  ]}
                />
              ) : (
                <Empty description="完成试运行后生成名单" />
              ),
            },
          ]}
        />
      </Card>
      {passed && (
        <div className="test-result">
          <CheckCircleFilled />
          <div>
            <b>功能验证与适配验证全部通过</b>
            <p>产品流程完整、字段映射有效、浙江地区参数已生效，无阻断问题。</p>
          </div>
          <span>测试报告：TEST-2026-0731-08</span>
        </div>
      )}
      <div className="pane-actions">
        <Button onClick={() => message.info("已创建整改事项")}>发现问题</Button>
        {passed && (
          <Button
            icon={<DownloadOutlined />}
            onClick={() => exportExcel(companyResults, "产品验证结果_演示.xls")}
          >
            导出Excel
          </Button>
        )}
        <Button type="primary" loading={running} onClick={passed ? next : run}>
          {passed ? "进入发布准备" : "运行完整流程"}
        </Button>
      </div>
    </div>
  );
}

function Publish({
  onPublish,
  next,
  go,
  onUse,
  productName,
  published,
}: {
  onPublish: () => void;
  next: () => void;
  go: (v: View) => void;
  onUse: (name: string) => void;
  productName: string;
  published: boolean;
}) {
  const [scope, setScope] = useState("全行发布");
  const [approving, setApproving] = useState(false);
  const [done, setDone] = useState(false);
  const [approvalStep, setApprovalStep] = useState(() =>
    typeof window !== "undefined"
      ? Number(localStorage.getItem("icbc-approval-step") || 0)
      : 0,
  );
  const tested =
    typeof window !== "undefined" &&
    localStorage.getItem("icbc-build-tested") === "true";
  const approvalNames = [
    "业务负责人确认",
    "数据管理部门审核",
    "科技开发确认",
    "产品发布审批",
  ];
  const submitApproval = () => {
    if (!tested) {
      message.warning("请先完成测试验证并取得通过结论");
      return;
    }
    if (approvalStep === 0) {
      setApprovalStep(1);
      localStorage.setItem("icbc-approval-step", "1");
      message.success("已提交审批，等待业务负责人确认");
    } else
      message.info(
        `当前等待“${approvalNames[approvalStep - 1]}”，请先完成该节点审核`,
      );
  };
  const saveApprovalLog = (mode: "逐级模拟" | "一键通过") =>
    localStorage.setItem(
      "icbc-approval-log",
      JSON.stringify(
        approvalNames.map((name, i) => ({
          node: name,
          result: "通过",
          operator: "当前总行产品管理员（演示代办）",
          order: i + 1,
          mode,
        })),
      ),
    );
  const passCurrent = () => {
    setApproving(true);
    setTimeout(() => {
      const nextStep = Math.min(5, approvalStep + 1);
      setApproving(false);
      setApprovalStep(nextStep);
      localStorage.setItem("icbc-approval-step", String(nextStep));
      if (nextStep === 5) {
        saveApprovalLog("逐级模拟");
        message.success("四级审批已全部通过，请确认发布正式产品");
      } else
        message.success(
          `${approvalNames[nextStep - 2]}已通过，已流转至下一审核节点`,
        );
    }, 500);
  };
  const passAll = () => {
    setApproving(true);
    setTimeout(() => {
      setApproving(false);
      setApprovalStep(5);
      localStorage.setItem("icbc-approval-step", "5");
      saveApprovalLog("一键通过");
      message.success("已按顺序模拟完成四级审批，请确认发布正式产品");
    }, 700);
  };
  const publishProduct = () => {
    if (approvalStep !== 5) {
      message.warning("四级审批全部通过后方可发布");
      return;
    }
    setDone(true);
    localStorage.setItem("icbc-published", "true");
    localStorage.setItem(
      "icbc-published-product",
      JSON.stringify({
        id: "P-2026-089",
        name: productName,
        scope: scope.replace("发布", ""),
        version: "V3.1",
        source: "某分行半导体产业链客户机会识别",
        sourceId: "CP-2026-0801-09",
      }),
    );
    onPublish();
    message.success("总行正式产品已发布，并同步至应用门户、成品区和来源分行");
  };
  if (done || published)
    return (
      <Result
        status="success"
        title="总行级正式产品发布成功"
        subTitle={`${productName} · V3.1 · 全行适用；某分行半导体原产品作为来源候选产品保留`}
        extra={
          <Space wrap>
            <Button onClick={next}>沉淀新增能力</Button>
            <Button
              onClick={() => {
                localStorage.setItem("icbc-open-published-detail", "true");
                go("finished");
              }}
            >
              查看正式产品详情
            </Button>
            <Button type="primary" onClick={() => onUse(productName)}>
              进入产品工作台运行
            </Button>
          </Space>
        }
      />
    );
  return (
    <div className="build-pane">
      <Row gutter={16}>
        <Col span={16}>
          <Card title="发布信息确认">
            <Descriptions
              bordered
              column={2}
              items={[
                { label: "总行正式产品名称", children: productName },
                { label: "拟发布编号", children: "P-2026-089" },
                {
                  label: "来源候选产品",
                  children: "某分行半导体产业链客户机会识别",
                },
                { label: "来源编号", children: "CP-2026-0801-09" },
                { label: "建设路径", children: "能力组合＋标准产品升级" },
                { label: "整体复用率", children: <b className="red">68%</b> },
                {
                  label: "测试结论",
                  children: (
                    <StatusTag status={tested ? "全部通过" : "待验证"} />
                  ),
                },
                { label: "维护主体", children: "总行数据管理部、公司金融部" },
                { label: "产品版本", children: "V3.1" },
                {
                  label: "产品归属",
                  children: <Tag color="red">总行级标准产品</Tag>,
                },
              ]}
            />
            <div className="identity-note">
              <SafetyCertificateOutlined />
              <span>
                <b>身份转换说明</b>
                某分行半导体产业链报送成果是产品来源和首批验证样本；本次发布的是经总行优化、全行适配和审批后的独立正式产品，维护责任、版本和适用范围均转为总行口径。
              </span>
            </div>
            <h3>选择发布类型</h3>
            <Radio.Group
              value={scope}
              onChange={(e) => setScope(e.target.value)}
              className="publish-scope"
            >
              <Radio.Button value="地方发布">
                地方发布
                <br />
                <small>仅四川省机构可见</small>
              </Radio.Button>
              <Radio.Button value="限域发布">
                限域发布
                <br />
                <small>指定机构与岗位</small>
              </Radio.Button>
              <Radio.Button value="全行发布">
                全行发布
                <br />
                <small>发布为总行级标准产品</small>
              </Radio.Button>
            </Radio.Group>
          </Card>
        </Col>
        <Col span={8}>
          <Card
            title="审批流程"
            extra={
              <Tag
                color={
                  approvalStep === 0
                    ? "default"
                    : approvalStep === 5
                      ? "green"
                      : "orange"
                }
              >
                {approvalStep === 0
                  ? "待提交"
                  : approvalStep === 5
                    ? "审批通过"
                    : "审核中"}
              </Tag>
            }
          >
            <Steps
              direction="vertical"
              size="small"
              current={Math.max(0, approvalStep - 1)}
              items={approvalNames.map((title, i) => ({
                title,
                description:
                  approvalStep === 0
                    ? "待提交审核"
                    : i < approvalStep - 1
                      ? "已通过 · 当前管理员演示代办"
                      : i === approvalStep - 1 && approvalStep < 5
                        ? "待当前审核"
                        : "待前序审核",
                status:
                  approvalStep > 0 && i < approvalStep - 1
                    ? "finish"
                    : approvalStep > 0 &&
                        i === approvalStep - 1 &&
                        approvalStep < 5
                      ? "process"
                      : approvalStep === 5
                        ? "finish"
                        : "wait",
              }))}
            />
            {approvalStep > 0 && approvalStep <= 4 && (
              <Space direction="vertical" style={{ width: "100%" }}>
                <Button
                  block
                  type="primary"
                  loading={approving}
                  onClick={passCurrent}
                >
                  模拟通过当前审核
                </Button>
                <Button block loading={approving} onClick={passAll}>
                  一键完成全部审批
                </Button>
              </Space>
            )}
            {approvalStep === 5 && (
              <Button block type="primary" onClick={publishProduct}>
                确认发布正式产品
              </Button>
            )}
          </Card>
        </Col>
      </Row>
      <div className="publish-note">
        <NotificationOutlined />
        <span>
          {tested
            ? approvalStep === 0
              ? "测试已通过，可提交审批；提交前所有节点均保持待审核状态。"
              : approvalStep === 5
                ? "四级审批已按顺序通过，产品尚未发布；点击“确认发布正式产品”后才会同步上架。"
                : "演示审批模式：当前总行管理员可逐级代办或一键完成全部审批，系统仍保留四级顺序与审批记录。"
            : "发布门禁尚未满足：请返回测试验证并完成全流程测试。"}
        </span>
      </div>
      <div className="pane-actions">
        <Button
          onClick={() =>
            message.info("已打开产品说明、测试报告、资源授权与组装清单预览")
          }
        >
          预览审批材料
        </Button>
        <Button
          type="primary"
          disabled={!tested || approvalStep > 0}
          onClick={submitApproval}
        >
          {approvalStep === 0
            ? "提交审批"
            : approvalStep === 5
              ? "审批已通过"
              : "审批处理中"}
        </Button>
      </div>
    </div>
  );
}

function Deposit({ onDone }: { onDone: () => void }) {
  const [done, setDone] = useState(false);
  if (done)
    return (
      <Result
        status="success"
        title="能力成果已回沉"
        subTitle="“产业事件影响评分”和“营销/风险双任务分派”已进入能力资产中心，可被其他事件类产品复用。"
        extra={
          <Button type="primary" onClick={onDone}>
            查看能力资产中心
          </Button>
        }
      />
    );
  return (
    <div className="build-pane">
      <div className="deposit-alert">
        <DeploymentUnitOutlined />
        <div>
          <b>检测到 2 项可沉淀成果</b>
          <p>
            本次建设新增事件影响评分和跨流程任务分派能力，可作为全行公共能力。
          </p>
        </div>
      </div>
      <Row gutter={16}>
        <Col span={12}>
          <Card className="deposit-card">
            <Checkbox defaultChecked />
            <Tag color="purple">能力模块</Tag>
            <h3>产业事件影响评分</h3>
            <p>
              综合事件类型、金额、时效、来源可信度、产业影响和客户关系，输出机会分与风险等级。
            </p>
            <Descriptions
              size="small"
              column={1}
              items={[
                {
                  label: "来源产品",
                  children: "产业事件驱动的融资机会与风险识别",
                },
                { label: "适用范围", children: "全行事件类产品" },
                { label: "维护主体", children: "总行数据管理部" },
              ]}
            />
          </Card>
        </Col>
        <Col span={12}>
          <Card className="deposit-card">
            <Checkbox defaultChecked />
            <Tag>流程能力</Tag>
            <h3>营销/风险双任务分派与反馈</h3>
            <p>
              依据事件信号类型、客户归属和管户关系分别生成营销或风险任务，并回收处置结果。
            </p>
            <Descriptions
              size="small"
              column={1}
              items={[
                {
                  label: "来源产品",
                  children: "产业事件驱动的融资机会与风险识别",
                },
                { label: "适用范围", children: "营销、风险与合规流程" },
                { label: "维护主体", children: "总行数据管理部、相关业务部门" },
              ]}
            />
          </Card>
        </Col>
      </Row>
      <Card title="成果登记信息">
        <Form layout="vertical">
          <Row gutter={16}>
            <Col span={8}>
              <Form.Item label="成果名称">
                <Input defaultValue="产业事件影响评分" />
              </Form.Item>
            </Col>
            <Col span={8}>
              <Form.Item label="成果类型">
                <Select
                  defaultValue="能力模块"
                  options={[
                    { value: "通用规则" },
                    { value: "能力模块" },
                    { value: "流程能力" },
                  ]}
                />
              </Form.Item>
            </Col>
            <Col span={8}>
              <Form.Item label="版本号">
                <Input defaultValue="V1.0" />
              </Form.Item>
            </Col>
            <Col span={12}>
              <Form.Item label="输入要求">
                <Input defaultValue="标准事件＋企业主体＋客户关系" />
              </Form.Item>
            </Col>
            <Col span={12}>
              <Form.Item label="输出结果">
                <Input defaultValue="机会评分＋风险等级＋建议动作" />
              </Form.Item>
            </Col>
          </Row>
        </Form>
      </Card>
      <div className="pane-actions">
        <Button>暂不回沉</Button>
        <Button
          type="primary"
          onClick={() => {
            setDone(true);
            localStorage.setItem("icbc-deposited", "true");
            message.success("新增成果已进入能力资产中心");
          }}
        >
          保存并回沉能力
        </Button>
      </div>
    </div>
  );
}

const agentSteps = [
  {
    title: "理解需求",
    desc: "抽取对象、任务、流程位置和预期动作",
    result: "已识别 6 类业务要素",
  },
  {
    title: "检索存量",
    desc: "比对产品、母版、能力、规则及历史方案",
    result: "发现 1 个母版、7 项能力",
  },
  {
    title: "设计方案",
    desc: "确定复用路径、模型边界和人工节点",
    result: "形成 68% 复用方案",
  },
  {
    title: "生成验证",
    desc: "构造样本、评价集、门禁指标和风险用例",
    result: "生成 18 条测试用例",
  },
  {
    title: "交付建设",
    desc: "写入流程、映射、规则、模型和发布材料",
    result: "等待人工确认写入",
  },
];

function ProductBuildAgent({
  go,
  embedded = false,
}: {
  go: (v: View) => void;
  embedded?: boolean;
}) {
  const [task, setTask] = useState(
    "将某分行半导体产业链客户机会识别方案升级为全行标准产品，持续识别政策、订单与价格事件，形成营销、授信和风险协同任务",
  );
  const [running, setRunning] = useState(false);
  const [step, setStep] = useState(4);
  const [accepted, setAccepted] = useState<string[]>([
    "需求结构化",
    "存量查重",
  ]);
  const [mode, setMode] = useState("审慎协同");
  const run = () => {
    setRunning(true);
    setStep(0);
    const timer = window.setInterval(
      () =>
        setStep((s) => {
          if (s >= 4) {
            window.clearInterval(timer);
            setRunning(false);
            message.success("建设分析已完成，已形成可审查方案");
            return 4;
          }
          return s + 1;
        }),
      420,
    );
  };
  const proposals = [
    {
      name: "需求结构化",
      type: "业务定义",
      confidence: 96,
      summary: "产业事件触发—链上主体识别—机会/风险评价—跨岗位任务—结果回流",
      target: "需求解析、场景映射",
    },
    {
      name: "存量查重",
      type: "复用判断",
      confidence: 92,
      summary: "复用先进制造母版、主体关联等7项能力；四川地方规则参数化保留",
      target: "查重比对、建设路径",
    },
    {
      name: "流程组装草案",
      type: "产品流程",
      confidence: 89,
      summary: "12个节点、3个人工门禁、2个业务系统回写接口",
      target: "流程编排、数据映射",
    },
    {
      name: "模型与策略建议",
      type: "决策能力",
      confidence: 84,
      summary: "事件影响模型V2.3＋平衡策略；价格风险短期权重下调5%",
      target: "行业模型、策略仿真",
    },
    {
      name: "验证与发布包",
      type: "质量门禁",
      confidence: 91,
      summary: "18条用例、5项业务指标、4类安全测试和灰度发布方案",
      target: "测试验证、发布准备",
    },
  ];
  const evidence = [
    ["候选产品方案", "某分行 · CP-2026-0801-09", "2026-08-01"],
    ["全行产品母版", "先进制造产业链协同经营 V3.0", "2026-07-22"],
    ["运行效果样本", "浙江、江苏、四川共 3,286 条任务", "2026-07-31"],
    ["治理规则", "总行外数产品发布门禁 V2.1", "2026-06-18"],
  ];
  return (
    <div className={embedded ? "agent-studio embedded" : "agent-studio"}>
      {!embedded && (
        <div className="page-heading">
          <div>
            <h1>
              产品建设辅助 <DemoTag />
            </h1>
            <p>
              以标准产品为边界，协助完成需求理解、存量复用、方案设计、验证发布和运行迭代
            </p>
          </div>
          <Space>
            <Button onClick={() => message.info("已打开建设分析记录")}>
              运行审计
            </Button>
            <Button type="primary" onClick={() => go("build")}>
              进入正式建设
            </Button>
          </Space>
        </div>
      )}
      <div className="agent-hero">
        <div className="agent-title">
          <span>
            <RobotOutlined />
          </span>
          <div>
            <Tag color="red">总行产品建设辅助 · V1.4</Tag>
            <h2>先形成可审查的建设方案，再由人员确认写入正式流程</h2>
            <p>
              只调用已授权资产；不直接发布产品，不替代业务、数据与科技审批。
            </p>
          </div>
        </div>
        <div className="agent-controls">
          <Select
            value={mode}
            onChange={setMode}
            options={["审慎协同", "标准协同", "探索分析"].map((value) => ({
              value,
            }))}
          />
          <Button
            type="primary"
            icon={<RobotOutlined />}
            loading={running}
            onClick={run}
          >
            {running ? "正在分析" : "重新运行分析"}
          </Button>
        </div>
      </div>
      <Card
        className="agent-task-card"
        title="本次建设任务"
        extra={<Tag color="green">权限校验通过</Tag>}
      >
        <Input.TextArea
          value={task}
          onChange={(e) => setTask(e.target.value)}
          autoSize={{ minRows: 2, maxRows: 4 }}
        />
        <div className="agent-scope">
          <span>
            <b>可调用资产</b> 86个产品 · 12个母版 · 31项能力 · 48条规则
          </span>
          <span>
            <b>禁止动作</b> 自动发布 · 修改锁定口径 · 直接写入生产系统
          </span>
          <span>
            <b>预算门限</b> 单次 ≤ 18,000 Token
          </span>
        </div>
      </Card>
      <div className="agent-layout">
        <Card title="建设分析过程" className="agent-trace">
          <Steps
            direction="vertical"
            current={step}
            items={agentSteps.map((x, i) => ({
              title: x.title,
              description: (
                <div>
                  <span>{x.desc}</span>
                  {i <= step && (
                    <small>
                      <CheckCircleFilled /> {x.result}
                    </small>
                  )}
                </div>
              ),
            }))}
          />
          <div className="agent-audit">
            <SafetyCertificateOutlined />
            <span>
              <b>全过程可追溯</b>
              保留任务、证据版本、工具调用、人工修改和最终写入记录。
            </span>
          </div>
        </Card>
        <div className="agent-main">
          <div className="agent-summary">
            <div>
              <small>建议建设路径</small>
              <b>母版适配＋能力复用</b>
            </div>
            <div>
              <small>预计整体复用率</small>
              <b>68%</b>
            </div>
            <div>
              <small>需人工确认</small>
              <b>{5 - accepted.length} 项</b>
            </div>
            <div>
              <small>预计建设周期</small>
              <b>12个工作日</b>
            </div>
          </div>
          <Card
            title="建设建议与人工确认"
            extra={<span className="muted">建议不会自动进入正式产品</span>}
          >
            <div className="proposal-list">
              {proposals.map((p) => {
                const done = accepted.includes(p.name);
                return (
                  <div
                    className={done ? "proposal accepted" : "proposal"}
                    key={p.name}
                  >
                    <div className="proposal-head">
                      <span>
                        <Tag color={done ? "green" : "orange"}>
                          {done ? "已确认" : "待确认"}
                        </Tag>
                        <b>{p.name}</b>
                        <small>
                          {p.type} · 置信度 {p.confidence}%
                        </small>
                      </span>
                      <Progress
                        type="circle"
                        size={42}
                        percent={p.confidence}
                        strokeColor={p.confidence >= 90 ? "#15945b" : "#d98700"}
                      />
                    </div>
                    <p>{p.summary}</p>
                    <div className="proposal-actions">
                      <span>确认后写入：{p.target}</span>
                      <Space>
                        <Button
                          size="small"
                          onClick={() =>
                            message.info("已展开证据、差异项和调用记录")
                          }
                        >
                          核对依据
                        </Button>
                        <Button
                          size="small"
                          type="primary"
                          disabled={done}
                          onClick={() => {
                            setAccepted((x) => [...x, p.name]);
                            message.success(`“${p.name}”已写入对应建设草稿`);
                          }}
                        >
                          {done ? "已写入草稿" : "确认并写入"}
                        </Button>
                      </Space>
                    </div>
                  </div>
                );
              })}
            </div>
          </Card>
        </div>
        <Card
          title="证据与治理"
          className="agent-evidence"
          extra={<Tag>{evidence.length}项</Tag>}
        >
          <List
            size="small"
            dataSource={evidence}
            renderItem={(x) => (
              <List.Item>
                <div>
                  <b>{x[0]}</b>
                  <span>{x[1]}</span>
                  <small>版本日期 {x[2]}</small>
                </div>
                <Button type="link" size="small">
                  查看
                </Button>
              </List.Item>
            )}
          />
          <h3>必须人工决策</h3>
          <div className="human-gates">
            {[
              "业务目标与适用范围",
              "总行锁定规则变更",
              "行内数据调用权限",
              "模型发布与灰度范围",
              "授信及风险任务推送",
            ].map((x, i) => (
              <span key={x}>
                <SafetyCertificateOutlined />
                <b>{x}</b>
                <small>{i < 2 ? "业务负责人" : "数据/科技/风险联合确认"}</small>
              </span>
            ))}
          </div>
          <Button
            block
            onClick={() => message.success("已生成建设分析报告与审计附件")}
          >
            导出建设分析报告
          </Button>
        </Card>
      </div>
      <div className="agent-handoff">
        <div>
          <CheckCircleFilled />
          <span>
            <b>交付的是可写入、可验证、可回退的建设草稿</b>
            <small>
              当前已确认 {accepted.length}/5 项；未确认内容不会进入正式流程。
            </small>
          </span>
        </div>
        <Space>
          <Button onClick={() => message.info("已保存为建设方案草稿 V0.6")}>
            保存方案
          </Button>
          <Button type="primary" onClick={() => go("build")}>
            进入建设中心逐项复核
          </Button>
        </Space>
      </div>
    </div>
  );
}

function SmartIntake({
  go,
  onProceed,
}: {
  go: (v: View) => void;
  onProceed: () => void;
}) {
  const [confirmed, setConfirmed] = useState([
    "业务目标与适用范围",
    "资源授权边界",
  ]);
  const decisions = [
    ["业务目标与适用范围", "业务负责人", "已确认"],
    ["资源授权边界", "数据管理部门", "已确认"],
    ["总行锁定规则变更", "业务负责人", "待确认"],
    ["模型与灰度发布范围", "数据/科技/风险联合确认", "待确认"],
  ];
  return (
    <div className="smart-intake">
      <div className="intake-banner">
        <div>
          <Tag color="red">智能受理结果</Tag>
          <h2>候选产品已完成机器预审，可以进入正式需求建模</h2>
          <p>
            本页展示受理结论与人工确认事项；完整依据和建设方案可在“产品建设辅助”中查看。
          </p>
        </div>
        <Space>
          <Button onClick={() => go("build-agent")}>查看完整建设分析</Button>
          <Button type="primary" onClick={onProceed}>
            进入需求建模
          </Button>
        </Space>
      </div>
      <div className="intake-summary">
        {[
          ["受理结论", "建议受理"],
          ["推荐路径", "母版适配＋能力复用"],
          ["存量相似度", "82%"],
          ["预计复用率", "68%"],
          ["待人工确认", `${decisions.length - confirmed.length}项`],
        ].map((x) => (
          <div key={x[0]}>
            <small>{x[0]}</small>
            <b>{x[1]}</b>
          </div>
        ))}
      </div>
      <Row gutter={[14, 14]}>
        <Col xs={24} xl={15}>
          <Card title="机器预审摘要" extra={<Tag color="green">预审完成</Tag>}>
            <Descriptions
              bordered
              size="small"
              column={2}
              items={[
                {
                  label: "业务任务",
                  children: "产业事件驱动的融资机会与风险识别",
                },
                { label: "适用对象", children: "先进制造产业链企业" },
                {
                  label: "主要输入",
                  children: "政策、订单、价格、工商、司法及行内授权数据",
                },
                {
                  label: "结果交付",
                  children: "机会名单、风险提示与跨岗位任务",
                },
                {
                  label: "复用判断",
                  children: "复用先进制造母版及7项公共能力",
                },
                {
                  label: "地方差异",
                  children: "四川产业目录与区域阈值参数化保留",
                },
              ]}
            />
            <div className="intake-route">
              <b>建议建设路径</b>
              <div>
                {[
                  "母版适配",
                  "能力匹配",
                  "地方参数配置",
                  "模型与策略验证",
                  "灰度发布",
                ].map((x, i) => (
                  <span key={x}>
                    <em>{i + 1}</em>
                    {x}
                  </span>
                ))}
              </div>
            </div>
          </Card>
        </Col>
        <Col xs={24} xl={9}>
          <Card
            title="受理门禁与人工确认"
            extra={
              <span className="muted">
                {confirmed.length}/{decisions.length} 已完成
              </span>
            }
          >
            <div className="intake-decisions">
              {decisions.map((x) => {
                const done = confirmed.includes(x[0]);
                return (
                  <div key={x[0]}>
                    <SafetyCertificateOutlined />
                    <span>
                      <b>{x[0]}</b>
                      <small>{x[1]}</small>
                    </span>
                    <Button
                      size="small"
                      type={done ? "default" : "primary"}
                      disabled={done}
                      onClick={() => {
                        setConfirmed((v) => [...v, x[0]]);
                        message.success(`已确认“${x[0]}”`);
                      }}
                    >
                      {done ? "已确认" : "确认"}
                    </Button>
                  </div>
                );
              })}
            </div>
          </Card>
        </Col>
      </Row>
      <Card title="写入正式建设流程" className="intake-handoff">
        <Steps
          size="small"
          current={0}
          items={[
            { title: "智能受理", description: "形成结构化受理单" },
            { title: "需求建模", description: "确认对象、任务与输出" },
            { title: "复用评估", description: "查重并选择建设路径" },
            { title: "产品组装", description: "配置流程、规则与模型" },
          ]}
        />
        <div>
          <span>
            受理单 V0.6 已保存，人工确认结果将随产品建设任务一并留痕。
          </span>
          <Button type="primary" onClick={onProceed}>
            确认受理并写入需求建模
          </Button>
        </div>
      </Card>
    </div>
  );
}

function ProductDefinition({ next }: { next: () => void }) {
  const fields = [
    ["目标对象", "先进制造业核心企业、供应商及项目参与企业"],
    ["业务问题", "重大订单与设备更新事件出现后，客户经理难以及时识别融资需求"],
    ["触发条件", "重大订单、中标、扩产、技改或项目开工事件达到可信度阈值"],
    ["数据依据", "政策、招投标、项目、工商、专精特新、司法与客户关系数据"],
    ["核心加工", "事件抽取—主体关联—影响判断—融资需求评分—风险核验"],
    ["产品输出", "机会名单、融资需求判断、证据链和客户经理任务"],
    ["流程位置", "营销线索发现—客户触达—融资需求记录—授信调查"],
    ["后续动作", "管户客户经理48小时内确认并记录触达、融资需求及转化状态"],
    ["成效指标", "命中率、按时处理率、融资需求转化率、单条有效线索成本"],
    ["适用边界", "先进制造业；地方口径与数据授权需经复制包适配"],
  ];
  const [confirmed, setConfirmed] = useState(fields.map(() => true));
  return (
    <div>
      <SectionTitle
        title="产品定义画布"
        sub="没有明确后续动作、成效指标和适用边界的需求，不进入标准产品建设。"
        extra={
          <Tag color={confirmed.every(Boolean) ? "green" : "orange"}>
            {confirmed.filter(Boolean).length}/10 已确认
          </Tag>
        }
      />
      <div className="definition-canvas">
        {fields.map((x, i) => (
          <Card
            size="small"
            key={x[0]}
            className={i === 1 || i === 7 || i === 8 ? "key-field" : ""}
            title={
              <>
                <span>0{i + 1}</span>
                {x[0]}
              </>
            }
            extra={
              <Checkbox
                checked={confirmed[i]}
                onChange={(e) =>
                  setConfirmed(
                    confirmed.map((v, j) => (j === i ? e.target.checked : v)),
                  )
                }
              >
                确认
              </Checkbox>
            }
          >
            <Input.TextArea
              autoSize={{ minRows: 2, maxRows: 4 }}
              defaultValue={x[1]}
            />
          </Card>
        ))}
      </div>
      <div className="pane-actions">
        <Button
          onClick={() => message.success("系统已补全画布并标注依据来源")}
          icon={<RobotOutlined />}
        >
          智能补全与证据核验
        </Button>
        <Button
          type="primary"
          disabled={!confirmed.every(Boolean)}
          onClick={next}
        >
          确认产品定义
        </Button>
      </div>
    </div>
  );
}

function CopyPackage({ next }: { next: () => void }) {
  const rows = [
    {
      asset: "产品母版",
      hq: "产业事件机会识别主流程",
      branch: "不可修改",
      version: "T3.2",
    },
    {
      asset: "标准能力",
      hq: "事件抽取、主体关联、司法核验",
      branch: "不可修改",
      version: "A2.7",
    },
    {
      asset: "行业与地区",
      hq: "标准行业代码、地域映射",
      branch: "可配置",
      version: "P1.4",
    },
    {
      asset: "时间与阈值",
      hq: "默认90日、机会分≥70",
      branch: "可配置",
      version: "S2.1",
    },
    {
      asset: "地方数据",
      hq: "字段与质量门槛",
      branch: "可替换映射",
      version: "M1.8",
    },
    {
      asset: "交付方式",
      hq: "平台、API、名单推送、任务接口",
      branch: "按授权选配",
      version: "D1.6",
    },
  ];
  return (
    <div>
      <SectionTitle
        title="产品复制包"
        sub="将标准产品连同资产清单、配置边界、验证用例、发布门禁和运营指标一并交付。"
        extra={<Tag color="red">COPY-PKG-089-V3.1</Tag>}
      />
      <Row gutter={[12, 12]}>
        <Col xs={24} xl={16}>
          <Card title="总行锁定项与分行可配置项">
            <Table
              rowKey="asset"
              pagination={false}
              dataSource={rows}
              columns={[
                { title: "复制对象", dataIndex: "asset" },
                { title: "标准内容", dataIndex: "hq" },
                {
                  title: "分行权限",
                  dataIndex: "branch",
                  render: (x: string) => (
                    <Tag color={x === "不可修改" ? "red" : "blue"}>{x}</Tag>
                  ),
                },
                { title: "绑定版本", dataIndex: "version" },
              ]}
            />
          </Card>
        </Col>
        <Col xs={24} xl={8}>
          <Card title="交付清单">
            <List
              dataSource={[
                "产品定义画布",
                "数据与字段映射模板",
                "规则/模型/策略版本",
                "历史样本与验证用例",
                "试点方案与发布门禁",
                "业务任务与回流字段",
                "质量边界与已知缺陷",
                "运行评价指标",
              ]}
              renderItem={(x, i) => (
                <List.Item>
                  <CheckCircleFilled className="green" /> {i + 1}. {x}
                </List.Item>
              )}
            />
            <Progress percent={100} strokeColor="#c7000b" />
            <Button
              block
              onClick={() => message.success("复制包已导出并生成适配说明")}
            >
              导出完整复制包
            </Button>
          </Card>
        </Col>
      </Row>
      <div className="pane-actions">
        <Button
          onClick={() => message.success("已模拟四川、江苏、北京三地参数适配")}
        >
          跨区域适配检查
        </Button>
        <Button type="primary" onClick={next}>
          确认复制边界并进入发布
        </Button>
      </div>
    </div>
  );
}

function CaseAssembly({ next }: { next: () => void }) {
  const [activeId, setActiveId] = useState("M001");
  const [showManifest, setShowManifest] = useState(false);
  const nodes = [
    { id: "IN", stage: "输入", name: "候选企业与外数资源", kind: "资源输入", input: "四川先进制造企业、工商、产业资质、招投标、司法数据", output: "待加工主体集", rule: "数据授权与时点校验", reuse: "标准资源 5项" },
    { id: "M001", stage: "01", name: "主体关联", kind: "能力模块", input: "企业名称、统一社会信用代码", output: "行内标准主体ID", rule: "主体匹配置信度≥0.90", reuse: "直接复用" },
    { id: "M002", stage: "02", name: "工商标准化", kind: "能力模块", input: "供应商工商原始字段", output: "标准企业信息", rule: "注册地=四川；成立≤近3年", reuse: "直接复用" },
    { id: "M008", stage: "03", name: "产业资质核验", kind: "能力模块", input: "主体ID、国家及地方资质名单", output: "专精特新/产业链资质标签", rule: "行业=先进制造；资质有效", reuse: "复用＋地方映射" },
    { id: "M007", stage: "04", name: "招投标活跃度计算", kind: "能力模块", input: "近12个月招投标记录", output: "活跃度评分、重大项目明细", rule: "观察期=12个月；活跃度≥60", reuse: "直接复用" },
    { id: "M005", stage: "05", name: "司法风险核验", kind: "能力模块", input: "主体ID、司法及经营异常记录", output: "风险标签与证据明细", rule: "近3个月重大风险=排除", reuse: "直接复用" },
    { id: "M004", stage: "06", name: "名单筛选与排序", kind: "能力模块", input: "资质、活跃度、风险标签", output: "候选企业名单与入选依据", rule: "全部准入规则＋机会分排序", reuse: "直接复用" },
    { id: "M020", stage: "07", name: "行动提示", kind: "能力模块", input: "企业画像、入选依据、客户归属", output: "客户经理跟进建议", rule: "按机构与岗位分派", reuse: "新增后回沉" },
    { id: "OUT", stage: "输出", name: "产业客户机会识别产品", kind: "标准产品", input: "完整加工链路", output: "名单、证据链、行动提示", rule: "地方试点→限域→全行", reuse: "待验证发布" },
  ];
  const active = nodes.find((x) => x.id === activeId) || nodes[1];
  const ruleRows = [
    { key: 1, rule: "地区筛选规则", param: "四川省", scope: "可配置", source: "R002" },
    { key: 2, rule: "行业筛选规则", param: "先进制造业", scope: "可配置", source: "R003" },
    { key: 3, rule: "注册年限规则", param: "近3年", scope: "可配置", source: "R001" },
    { key: 4, rule: "招投标活跃度规则", param: "近12个月 / ≥60分", scope: "可配置", source: "R007" },
    { key: 5, rule: "司法风险排除规则", param: "近3个月重大风险排除", scope: "全行锁定", source: "R006" },
    { key: 6, rule: "四川产业目录映射", param: "地方目录→标准行业代码", scope: "本地扩展", source: "LOCAL-01" },
  ];
  return (
    <div className="build-pane case-assembly">
      <div className="assembly-hero">
        <div>
          <Tag color="red">报告案例 · 模块化组装</Tag>
          <h2>某分行半导体产业链客户机会识别 → 全行可复用标准产品</h2>
          <p>以成熟产品母版为骨架，逐项挂接能力模块、通用规则、数据映射和地方扩展；点击任一节点查看输入、输出与配置边界。</p>
        </div>
        <div className="assembly-kpis">
          <span><b>7</b><small>能力模块</small></span>
          <span><b>6</b><small>规则参数</small></span>
          <span><b>68%</b><small>预计复用率</small></span>
          <span><b>12天</b><small>预计周期</small></span>
        </div>
      </div>
      <div className="asset-formula">
        <span><small>产品母版</small><b>产业客户机会识别母版</b><Tag color="green">复用骨架</Tag></span><em>＋</em>
        <span><small>能力模块</small><b>7个独立加工环节</b><Tag color="green">按契约调用</Tag></span><em>＋</em>
        <span><small>通用规则</small><b>地区 · 行业 · 期限 · 阈值</b><Tag color="blue">参数化</Tag></span><em>＋</em>
        <span><small>本地扩展</small><b>四川产业目录及政策字段</b><Tag color="orange">隔离保留</Tag></span>
      </div>
      <Row gutter={14} align="stretch">
        <Col xs={24} xl={17}>
          <Card title="可视化产品组装画布" extra={<Tag color="processing">输入输出契约：9/9 通过</Tag>} className="assembly-canvas-card">
            <div className="assembly-lanes">
              <div className="lane-labels"><span>资源输入</span><span>标准加工链</span><span>业务交付</span></div>
              <div className="assembly-chain">
                {nodes.map((node, i) => <React.Fragment key={node.id}>
                  <button className={`${activeId === node.id ? "active" : ""} ${node.kind === "能力模块" ? "module" : node.kind === "标准产品" ? "product" : "endpoint"}`} onClick={() => setActiveId(node.id)}>
                    <small>{node.stage} · {node.id}</small><b>{node.name}</b><em>{node.reuse}</em>
                  </button>{i < nodes.length - 1 && <i>→</i>}
                </React.Fragment>)}
              </div>
            </div>
            <div className="assembly-legend"><span className="module-dot" />能力模块 <span className="rule-dot" />通用规则挂接 <span className="local-dot" />本地扩展 <span className="product-dot" />标准产品输出</div>
          </Card>
        </Col>
        <Col xs={24} xl={7}>
          <Card title="节点说明" className="node-inspector">
            <Tag color={active.kind === "能力模块" ? "red" : active.kind === "标准产品" ? "green" : "blue"}>{active.kind}</Tag>
            <h3>{active.name}</h3><p>{active.id} · {active.stage}</p>
            <dl><dt>标准输入</dt><dd>{active.input}</dd><dt>处理输出</dt><dd>{active.output}</dd><dt>挂接规则</dt><dd>{active.rule}</dd><dt>复用判断</dt><dd>{active.reuse}</dd></dl>
            <Button block onClick={() => message.info(`${active.name}：输入输出契约、依赖版本与异常处理已展开`)}>查看模块契约</Button>
          </Card>
        </Col>
      </Row>
      <Row gutter={14} style={{ marginTop: 14 }}>
        <Col xs={24} xl={15}><Card title="规则参数与地方扩展"><Table size="small" pagination={false} rowKey="key" dataSource={ruleRows} columns={[{title:"规则",dataIndex:"rule"},{title:"本产品取值",dataIndex:"param"},{title:"配置边界",dataIndex:"scope",render:(x:string)=><Tag color={x==="可配置"?"blue":x==="本地扩展"?"orange":"red"}>{x}</Tag>},{title:"资产编号",dataIndex:"source"}]} /></Card></Col>
        <Col xs={24} xl={9}><Card title="组装完成后自动形成"><div className="assembly-output-list"><span><CheckCircleFilled /><b>产品组装清单</b><small>1个母版＋7个模块＋6条规则＋12项映射</small></span><span><CheckCircleFilled /><b>全链路血缘</b><small>每项结果可追溯至资源、模块和规则版本</small></span><span><CheckCircleFilled /><b>能力回沉候选</b><small>行动提示模块＋四川产业目录映射规则</small></span><span><CheckCircleFilled /><b>三类验证用例</b><small>功能、数据适配与运行验证</small></span></div></Card></Col>
      </Row>
      {showManifest && <Card title="产品组装清单 · ASSEMBLY-V1.0" style={{ marginTop: 14 }}><Descriptions bordered size="small" column={2} items={[{label:"母版",children:"T004 产业客户机会识别母版 V3.1"},{label:"标准资源",children:"工商、资质、招投标、司法等5项"},{label:"能力模块",children:"M001/M002/M008/M007/M005/M004/M020"},{label:"通用规则",children:"R001/R002/R003/R006/R007＋LOCAL-01"},{label:"字段映射",children:"12项，全部通过语义与类型校验"},{label:"本地扩展",children:"四川产业目录、地方政策字段"},{label:"输出成果",children:"候选名单、入选依据、行动提示",span:2}]} /></Card>}
      <div className="pane-actions"><Button onClick={() => setShowManifest(!showManifest)}>{showManifest ? "收起组装清单" : "生成产品组装清单"}</Button><Button type="primary" onClick={next}>锁定组装版本，进入测试验证</Button></div>
    </div>
  );
}

function StandardBuildStep({
  stage,
  next,
  published,
  onPublish,
  onUse,
  productName,
}: {
  stage: string;
  next: () => void;
  published: boolean;
  onPublish: () => void;
  onUse: (name: string) => void;
  productName: string;
}) {
  const [verified, setVerified] = useState(false);
  if (stage === "intake")
    return <div className="build-pane simplified-build"><div className="build-conclusion"><RobotOutlined /><div><Tag color="green">智能识别完成</Tag><h2>该地方成果与全行存量产品部分重合，建议进入标准化改造</h2><p>系统已把报送材料转换为统一产品描述，并同步完成场景映射和存量查重。</p></div><strong>76%<small>最高相似度</small></strong></div><Row gutter={14}><Col span={14}><Card title="统一产品定义"><Descriptions bordered size="small" column={2} items={[{label:"业务任务",children:"产业事件驱动的融资机会与风险识别"},{label:"目标对象",children:"先进制造产业链企业"},{label:"主要输入",children:"政策、订单、价格、工商及司法数据"},{label:"结果交付",children:"机会与风险信号、证据链、协同任务"},{label:"流程位置",children:"营销拓展—授信核验—风险监测"},{label:"适用边界",children:"四川地方口径需参数化处理"}]} /></Card></Col><Col span={10}><Card title="识别结论"><div className="simple-decision"><span><small>存量对照</small><b>先进制造产业链协同经营 V3.0</b></span><span><small>建设路径</small><b>母版适配＋能力复用</b></span><span><small>本阶段产物</small><b>产品定义书 REQ-V1.0</b></span></div></Card></Col></Row><div className="included-work">已纳入：需求受理、产品定义、场景映射、存量查重、路径选择</div><div className="pane-actions"><Button type="primary" onClick={next}>确认定义，进入拆解复用</Button></div></div>;
  if (stage === "reuse") {
    const rows=[{key:"1",part:"产品母版",decision:"复用",content:"先进制造产业链协同经营母版",reason:"业务链路与输出结构一致"},{key:"2",part:"能力模块",decision:"复用",content:"主体关联、事件识别、司法核验等7项",reason:"输入输出契约可直接调用"},{key:"3",part:"通用规则",decision:"参数化",content:"行业、地域、期限、风险阈值",reason:"差异可转换为配置参数"},{key:"4",part:"地方内容",decision:"保留",content:"四川产业目录与地方政策口径",reason:"依赖地方数据和管理要求"}];
    return <div className="build-pane simplified-build"><div className="build-conclusion"><ApartmentOutlined /><div><Tag color="red">标准化不是复制产品，而是重构产品</Tag><h2>将地方成果拆成可复用、可参数化和需本地保留的内容</h2><p>比较不再停留在产品名称，而是深入业务任务、数据输入、加工逻辑、结果交付和后续动作。</p></div><strong>68%<small>预计复用率</small></strong></div><Card title="结构化拆解与复用判断"><Table size="small" pagination={false} dataSource={rows} columns={[{title:"拆解对象",dataIndex:"part"},{title:"标准化判断",dataIndex:"decision",render:(x)=><Tag color={x==="复用"?"green":x==="参数化"?"blue":"orange"}>{x}</Tag>},{title:"具体内容",dataIndex:"content"},{title:"判断依据",dataIndex:"reason"}]} /></Card><div className="simple-output-row"><span><small>直接复用</small><b>1个母版＋7项能力</b></span><span><small>参数化</small><b>6条规则</b></span><span><small>数据映射</small><b>12个字段</b></span><span><small>本地扩展</small><b>2项</b></span></div><div className="included-work">已纳入：结构化拆解、建设路径、能力匹配、规则判断、地方边界确认</div><div className="pane-actions"><Button type="primary" onClick={next}>确认复用方案，进入模块组装</Button></div></div>;
  }
  if (stage === "assembly")
    return <div className="build-pane simplified-build"><div className="build-conclusion"><BuildOutlined /><div><Tag color="blue">模块化组装</Tag><h2>以产品母版为骨架，组合能力、规则、数据映射和地方参数</h2><p>系统已生成全行标准产品草案；技术细节收纳在各模块中，不再拆成多个独立页面。</p></div><strong>12天<small>预计周期</small></strong></div><Card title="标准产品组装方案"><div className="simple-assembly"><div><small>产品母版</small><b>先进制造产业链协同经营</b><Tag color="green">整体骨架</Tag></div><em>＋</em><div><small>能力模块</small><b>事件识别 · 主体关联 · 影响判断 · 风险核验</b><Tag color="green">复用7项</Tag></div><em>＋</em><div><small>通用规则</small><b>行业 · 地域 · 期限 · 阈值 · 排除条件</b><Tag color="blue">参数化6条</Tag></div><em>＋</em><div><small>地方扩展</small><b>四川目录 · 地方政策字段</b><Tag color="orange">保留2项</Tag></div></div></Card><Row gutter={14}><Col span={12}><Card title="统一业务链路"><div className="simple-flow">{["事件发现","主体关联","机会/风险判断","任务推送","反馈回流"].map((x,i)=><React.Fragment key={x}><span>{x}</span>{i<4&&<em>→</em>}</React.Fragment>)}</div></Card></Col><Col span={12}><Card title="组装检查"><div className="simple-checks"><span><CheckCircleFilled /> 数据字段已映射 28/28</span><span><CheckCircleFilled /> 模块输入输出契约通过</span><span><CheckCircleFilled /> 地方参数与全行口径已隔离</span><span><CheckCircleFilled /> 结果血缘与责任部门已绑定</span></div></Card></Col></Row><div className="included-work">已纳入：数据映射、能力与规则配置、流程编排、接口绑定、版本记录</div><div className="pane-actions"><Button type="primary" onClick={next}>生成标准产品，进入验证发布</Button></div></div>;
  return <div className="build-pane simplified-build"><div className="build-conclusion"><SafetyCertificateOutlined /><div><Tag color={published?"green":"orange"}>{published?"正式产品已发布":"待验证发布"}</Tag><h2>用代表性样本验证可用性，发布后同步沉淀新增能力与规则</h2><p>验证、审批、发布和能力回沉集中在同一页完成，形成“建设—应用—反馈—迭代”闭环。</p></div><strong>{published?"V1.0":"3项"}<small>{published?"正式版本":"发布门禁"}</small></strong></div><Row gutter={14}><Col span={15}><Card title="验证与发布门禁"><div className="release-gates">{[["功能验证","关键流程及输出正确","通过"],["适配验证","四川、江苏、北京样本运行","通过"],["业务验证","营销、授信、风险任务可回流",verified?"通过":"待执行"]].map(x=><div key={x[0]}><CheckCircleFilled className={x[2]==="通过"?"green":"muted"}/><span><b>{x[0]}</b><small>{x[1]}</small></span><Tag color={x[2]==="通过"?"green":"orange"}>{x[2]}</Tag></div>)}</div>{!verified&&<Button block type="primary" onClick={()=>{setVerified(true);message.success("业务验证完成，全部发布门禁已通过")}}>运行代表性样本验证</Button>}</Card></Col><Col span={9}><Card title="发布后自动形成"><div className="simple-decision"><span><small>正式产品</small><b>全国产业链机会与风险识别 V1.0</b></span><span><small>能力回沉</small><b>事件影响判断能力 1项</b></span><span><small>规则回沉</small><b>地方产业目录映射规则 2条</b></span></div></Card></Col></Row><div className="included-work">已纳入：功能测试、适配验证、业务试点、审批发布、版本定版、评价指标与能力回沉</div><div className="pane-actions">{published?<Button type="primary" onClick={()=>onUse(productName)}>查看正式产品可视化</Button>:<Button type="primary" disabled={!verified} onClick={()=>{onPublish();message.success("标准产品已发布，并完成能力与规则回沉")}}>确认发布正式产品</Button>}</div></div>;
}

function BuildCenter({
  published,
  onPublish,
  onReset,
  go,
  onUse,
  productName,
  initialTab = "agent",
}: {
  published: boolean;
  onPublish: () => void;
  onReset: () => void;
  go: (v: View) => void;
  onUse: (name: string) => void;
  productName: string;
  initialTab?: string;
}) {
  const [tab, setTab] = useState(() => {
    if (initialTab !== "agent") return initialTab;
    if (typeof window !== "undefined" && localStorage.getItem("icbc-requirement-baseline")) return "dedupe";
    return "agent";
  });
  const [completedTabs, setCompletedTabs] = useState<string[]>(() => {
    if (typeof window === "undefined") return [];
    return JSON.parse(localStorage.getItem("icbc-build-completed-tabs") || "[]");
  });
  const [candidateId, setCandidateId] = useState(candidates[0].id);
  const [resetKey, setResetKey] = useState(0);
  const selectedCandidate = candidates.find((x) => x.id === candidateId)!;
  const profile = buildProfiles[candidateId];
  const idx = buildTabs.findIndex((x) => x[0] === tab);
  const next = () => {
    const completed = [...new Set([...completedTabs, tab])];
    setCompletedTabs(completed);
    localStorage.setItem("icbc-build-completed-tabs", JSON.stringify(completed));
    const nextTab = buildTabs[Math.min(buildTabs.length - 1, idx + 1)][0];
    if (tab === "scene") {
      localStorage.setItem("icbc-requirement-baseline", "REQ-V1.0");
      message.success("需求受理与定义已完成，REQ-V1.0 基线已锁定，现进入复用评估");
    }
    setTab(nextTab);
  };
  const resetBuild = () => {
    localStorage.removeItem("icbc-published");
    localStorage.removeItem("icbc-deposited");
    localStorage.removeItem("icbc-build-tested");
    localStorage.removeItem("icbc-approval-step");
    localStorage.removeItem("icbc-approval-log");
    localStorage.removeItem("icbc-published-product");
    localStorage.removeItem("icbc-build-completed-tabs");
    localStorage.removeItem("icbc-requirement-baseline");
    setCompletedTabs([]);
    setCandidateId(candidates[0].id);
    setTab("agent");
    setResetKey((x) => x + 1);
    onReset();
    message.success("已恢复默认候选产品，并从受理校验重新开始");
  };
  const render = () =>
    ({
      agent: <SmartIntake go={go} onProceed={next} />,
      definition: <ProductDefinition next={next} />,
      scene: <SceneMapping next={next} profile={profile} />,
      dedupe: <Dedupe next={next} go={go} />,
      path: <PathChoice next={next} />,
      breakdown: <Breakdown next={next} />,
      mapping: <DataMapping next={next} />,
      match: <CapabilityMatch next={next} />,
      canvas: <FlowCanvas next={next} profile={profile} />,
      testing: <Testing next={next} />,
      simulation: <StrategySimulation next={next} />,
      version: <ModelVersion next={next} />,
      publish: <Publish onPublish={onPublish} next={next} go={go} onUse={onUse} productName={productName} published={published} />,
      attribution: <EffectAttribution next={next} />,
      deposit: <Deposit onDone={() => go("assets")} />,
    })[tab];
  const currentStageIndex = Math.max(
    0,
    buildStages.findIndex((stage) => stage.tabs.includes(tab)),
  );
  const currentStage = buildStages[currentStageIndex];
  const stageDone = (stage: any) =>
    stage.tabs.every((key: string) => completedTabs.includes(key));
  const status =
    published && candidateId === candidates[0].id
      ? "已转正式产品"
      : currentStageIndex === 0
        ? "需求定义中"
        : currentStageIndex < 3
          ? "建设中"
          : currentStageIndex === 3
            ? "待验证"
            : currentStageIndex === 4
              ? "待发布"
              : "运营中";
  const progress =
    published && candidateId === candidates[0].id
      ? 100
      : Math.round((completedTabs.length / buildTabs.length) * 100);
  const switchCandidate = (id: string) => {
    setCandidateId(id);
    setTab("agent");
    setResetKey((x) => x + 1);
    message.info("已切换候选产品，并重新生成建设计划");
  };
  const candidateMenu = {
    items: candidates.map((x) => ({
      key: x.id,
      label: (
        <div className="candidate-option">
          <b>{x.name}</b>
          <small>
            {x.id} · {x.org} · {x.status}
          </small>
        </div>
      ),
      onClick: () => switchCandidate(x.id),
    })),
  };
  return (
    <div className="build-center">
      <div className="build-fixed-head">
        <div className="product-identity">
          <Tag color="red">候选产品</Tag>
          <div>
            <div className="product-title-line">
              <h1>{selectedCandidate.name}</h1>
              <Dropdown
                menu={candidateMenu}
                trigger={["click"]}
                overlayClassName="candidate-dropdown"
              >
                <Button
                  className="candidate-switcher"
                  size="small"
                  aria-label="切换待建设产品"
                  title="切换待建设产品"
                >
                  <SwapOutlined />
                </Button>
              </Dropdown>
            </div>
            <p>
              {selectedCandidate.id} · {selectedCandidate.org} · 负责人：
              {selectedCandidate.owner}
            </p>
          </div>
        </div>
        <div className="build-progress">
          <span>
            <small>当前状态</small>
            <StatusTag status={status} />
          </span>
          <span className="progress-item">
            <small>建设进度</small>
            <Progress percent={progress} size="small" />
          </span>
          <span>
            <small>预计复用</small>
            <b>{selectedCandidate.score}%</b>
          </span>
        </div>
        <Space className="build-head-actions">
          <Tag color="red" icon={<BuildOutlined />}>建设流程已加载</Tag>
          <Button icon={<ReloadOutlined />} onClick={resetBuild}>
            重置
          </Button>
          <Button onClick={() => message.success("建设草稿已保存")}>
            保存
          </Button>
          <Button
            type="primary"
            onClick={() => message.success("已提交当前阶段评审")}
          >
            提交评审
          </Button>
        </Space>
      </div>
      <div className="build-stages">
        {buildStages.map((stage, i) => (
          <button
            key={stage.name}
            className={`${currentStageIndex === i ? "active" : stageDone(stage) ? "done" : ""} ${stage.focus ? "focus-stage" : ""}`}
            onClick={() => setTab(stage.tabs[0])}
          >
            <span>{stageDone(stage) ? <CheckCircleFilled /> : i + 1}</span>
            <b>{stage.name}</b>
            <small>{stage.desc}{stage.focus ? " · 汇报重点" : ""}</small>
          </button>
        ))}
      </div>
      <div className="stage-workbar">
        <div className="stage-current">
          <b>
            {currentStageIndex + 1}. {currentStage.name}
          </b>
          <span>{currentStage.desc}</span>
        </div>
        <div className="stage-subtabs">
          {currentStage.tabs.map((tabKey) => {
            const item = buildTabs.find((x) => x[0] === tabKey)!;
            return (
              <button
                key={tabKey}
                className={
                  tab === tabKey ? "active" : completedTabs.includes(tabKey) ? "done" : ""
                }
                onClick={() => setTab(tabKey)}
              >
                {completedTabs.includes(tabKey) && <CheckCircleFilled />}
                {item[1]}
              </button>
            );
          })}
        </div>
        <div className="stage-status">
          <span>
            <CheckCircleFilled /> 资源字段 28/32
          </span>
          <button onClick={() => go("procurement")}>
            <NotificationOutlined /> 采购缺口 4
          </button>
          <span>
            <SafetyCertificateOutlined /> 门禁：测试＋审批
          </span>
        </div>
      </div>
      <React.Fragment key={`${candidateId}-${resetKey}-${tab}`}>
        {render()}
      </React.Fragment>
    </div>
  );
}

const firstPageResourceProfiles: Record<string, any> = {
  "RES-JUD-2026-081": {
    summary:
      "按行内统一主体和风险口径提供企业司法事件明细，支持风险核验、分级预警与名单排除。",
    purposes: ["授信核验", "风险监测", "名单排除"],
    products: ["企业司法风险监测", "授信尽调外部信息核验"],
    params: [
      { name: "enterprise_id", cn: "企业主体ID", value: "ENT-110108-003821" },
      { name: "months", cn: "回溯月数", value: "3" },
    ],
    fields: [
      [
        "enterprise_id",
        "企业主体ID",
        "行内统一识别企业的主体编号",
        "string",
        "ENT-110108-003821",
      ],
      [
        "risk_event_type",
        "风险事件类型",
        "诉讼、执行、失信等统一事件分类",
        "string",
        "被执行",
      ],
      [
        "risk_level",
        "风险等级",
        "按行内规则计算的高、中、低风险",
        "string",
        "高",
      ],
      [
        "event_date",
        "事件日期",
        "司法事件公开或发生日期",
        "date",
        "2026-07-18",
      ],
      [
        "case_amount",
        "涉案金额（元）",
        "案件涉及金额，统一换算为人民币元",
        "decimal",
        "1265000",
      ],
      [
        "source_record_id",
        "来源记录ID",
        "用于追溯供应商原始记录",
        "string",
        "SRC-20260718-081",
      ],
    ],
    rows: [
      [
        "ENT-110108-003821",
        "被执行",
        "高",
        "2026-07-18",
        1265000,
        "SRC-20260718-081",
      ],
      [
        "ENT-110114-002176",
        "裁判文书",
        "中",
        "2026-07-12",
        380000,
        "SRC-20260712-094",
      ],
      [
        "ENT-110105-008923",
        "经营异常",
        "低",
        "2026-07-09",
        0,
        "SRC-20260709-117",
      ],
    ],
  },
  "EXT-R001": {
    summary:
      "汇集全国企业登记照面与存续状态，用于企业主体识别、基本情况核验和客户画像。",
    purposes: ["主体识别", "工商核验", "企业画像"],
    products: ["专精特新企业营销名单", "授信尽调外部信息核验"],
    params: [
      {
        name: "credit_code",
        cn: "统一社会信用代码",
        value: "91110108MA01X8KQ2N",
      },
      { name: "include_history", cn: "是否返回历史变更", value: "false" },
    ],
    fields: [
      [
        "enterprise_id",
        "企业主体ID",
        "行内统一企业主体编号",
        "string",
        "ENT-110108-003821",
      ],
      [
        "enterprise_name",
        "企业名称",
        "市场监管部门登记的企业全称",
        "string",
        "北京智造科技有限公司",
      ],
      [
        "credit_code",
        "统一社会信用代码",
        "企业法定身份识别代码",
        "string",
        "91110108MA01X8KQ2N",
      ],
      [
        "legal_representative",
        "法定代表人",
        "登记在册的法定代表人姓名",
        "string",
        "张明远",
      ],
      [
        "registered_capital",
        "注册资本（万元）",
        "企业登记注册资本，统一为万元",
        "decimal",
        "5000",
      ],
      [
        "established_date",
        "成立日期",
        "企业依法成立日期",
        "date",
        "2024-03-18",
      ],
      [
        "business_status",
        "登记状态",
        "存续、注销、吊销等登记状态",
        "string",
        "存续",
      ],
      [
        "industry_name",
        "所属行业",
        "按国民经济行业分类映射的行业名称",
        "string",
        "通用设备制造业",
      ],
    ],
    rows: [
      [
        "ENT-110108-003821",
        "北京智造科技有限公司",
        "91110108MA01X8KQ2N",
        "张明远",
        5000,
        "2024-03-18",
        "存续",
        "通用设备制造业",
      ],
      [
        "ENT-320115-006214",
        "南京芯联装备股份有限公司",
        "91320115MA27Q8L61P",
        "周启航",
        8200,
        "2022-09-06",
        "存续",
        "专用设备制造业",
      ],
      [
        "ENT-330106-004588",
        "杭州云控科技有限公司",
        "91330106MA2J7P3R8M",
        "陈嘉禾",
        3000,
        "2023-06-12",
        "存续",
        "软件和信息技术服务业",
      ],
    ],
  },
  "EXT-R002": {
    summary:
      "穿透展示企业股东、持股比例、最终受益人与实际控制关系，用于集团客户识别和关联风险分析。",
    purposes: ["股权穿透", "实控人识别", "集团客户识别"],
    products: ["集团客户关联关系识别", "授信尽调外部信息核验"],
    params: [
      {
        name: "enterprise_name",
        cn: "企业名称",
        value: "北京智造科技有限公司",
      },
      { name: "penetration_level", cn: "穿透层级", value: "3" },
    ],
    fields: [
      [
        "enterprise_name",
        "企业名称",
        "被查询企业的登记全称",
        "string",
        "北京智造科技有限公司",
      ],
      [
        "shareholder_name",
        "股东名称",
        "直接或间接股东名称",
        "string",
        "北京先进制造产业基金",
      ],
      [
        "shareholder_type",
        "股东类型",
        "企业、自然人或其他组织",
        "string",
        "企业法人",
      ],
      [
        "holding_ratio",
        "持股比例（%）",
        "该层股东对目标企业的持股比例",
        "decimal",
        "36.50",
      ],
      [
        "control_path",
        "控制路径",
        "从最终控制人至目标企业的股权链路",
        "string",
        "产业基金→智造控股→目标企业",
      ],
      [
        "ultimate_controller",
        "实际控制人",
        "股权穿透后识别的最终控制主体",
        "string",
        "北京市国资委",
      ],
      [
        "relation_level",
        "关系层级",
        "当前股权关系所处穿透层级",
        "integer",
        "2",
      ],
    ],
    rows: [
      [
        "北京智造科技有限公司",
        "北京先进制造产业基金",
        "企业法人",
        36.5,
        "产业基金→智造控股→目标企业",
        "北京市国资委",
        2,
      ],
      [
        "北京智造科技有限公司",
        "智造控股有限公司",
        "企业法人",
        51,
        "智造控股→目标企业",
        "北京市国资委",
        1,
      ],
      [
        "北京智造科技有限公司",
        "张明远",
        "自然人",
        12.5,
        "张明远→目标企业",
        "张明远",
        1,
      ],
    ],
  },
  "EXT-R003": {
    summary:
      "归集全国招标、中标和采购公告，支持识别企业经营活跃度、项目机会与潜在融资需求。",
    purposes: ["商机发现", "经营活跃度", "项目融资跟踪"],
    products: ["招投标客户发现", "产业事件驱动的融资机会识别"],
    params: [
      { name: "keyword", cn: "项目关键词", value: "智能制造" },
      { name: "region", cn: "项目地区", value: "北京市" },
    ],
    fields: [
      [
        "notice_id",
        "公告编号",
        "招投标公告的唯一标识",
        "string",
        "BID-20260728-0186",
      ],
      [
        "project_name",
        "项目名称",
        "招标或中标项目名称",
        "string",
        "智能制造产线升级项目",
      ],
      [
        "notice_type",
        "公告类型",
        "招标、中标、变更或终止公告",
        "string",
        "中标公告",
      ],
      [
        "purchaser_name",
        "采购人名称",
        "项目采购或招标主体",
        "string",
        "北京高端装备产业园",
      ],
      [
        "winner_name",
        "中标企业",
        "公告披露的中标主体",
        "string",
        "北京智造科技有限公司",
      ],
      [
        "project_amount",
        "项目金额（万元）",
        "公告披露的项目含税金额",
        "decimal",
        "2860",
      ],
      ["publish_date", "公告日期", "公告公开发布日期", "date", "2026-07-28"],
      ["project_region", "项目地区", "项目实施所在地", "string", "北京市"],
    ],
    rows: [
      [
        "BID-20260728-0186",
        "智能制造产线升级项目",
        "中标公告",
        "北京高端装备产业园",
        "北京智造科技有限公司",
        2860,
        "2026-07-28",
        "北京市",
      ],
      [
        "BID-20260726-0912",
        "新能源电池检测设备采购",
        "招标公告",
        "华北新能源研究院",
        "—",
        1350,
        "2026-07-26",
        "天津市",
      ],
      [
        "BID-20260724-0437",
        "工业机器人柔性产线建设",
        "中标公告",
        "南京先进制造有限公司",
        "南京芯联装备股份有限公司",
        4180,
        "2026-07-24",
        "江苏省",
      ],
    ],
  },
  "EXT-R004": {
    summary:
      "提供国家级和省级专精特新企业认定名单、批次与有效期，用于资质核验和营销名单筛选。",
    purposes: ["资质核验", "客群筛选", "政策匹配"],
    products: ["专精特新企业营销名单", "区域产业链客户画像"],
    params: [
      { name: "region", cn: "认定地区", value: "浙江省" },
      { name: "qualification_level", cn: "认定级别", value: "专精特新小巨人" },
    ],
    fields: [
      [
        "enterprise_name",
        "企业名称",
        "获得认定的企业登记全称",
        "string",
        "杭州云控科技有限公司",
      ],
      [
        "credit_code",
        "统一社会信用代码",
        "用于与工商主体进行精确关联",
        "string",
        "91330106MA2J7P3R8M",
      ],
      [
        "qualification_level",
        "认定级别",
        "创新型中小企业、专精特新或小巨人",
        "string",
        "专精特新小巨人",
      ],
      ["batch_name", "认定批次", "主管部门公布的认定批次", "string", "第六批"],
      ["recognized_date", "认定日期", "名单正式公布日期", "date", "2024-09-02"],
      [
        "valid_until",
        "有效期至",
        "当前资质有效期截止日期",
        "date",
        "2027-09-01",
      ],
      [
        "recognized_region",
        "认定地区",
        "作出认定的省级行政区",
        "string",
        "浙江省",
      ],
      [
        "competent_authority",
        "认定部门",
        "发布名单的主管部门",
        "string",
        "工业和信息化部",
      ],
    ],
    rows: [
      [
        "杭州云控科技有限公司",
        "91330106MA2J7P3R8M",
        "专精特新小巨人",
        "第六批",
        "2024-09-02",
        "2027-09-01",
        "浙江省",
        "工业和信息化部",
      ],
      [
        "宁波精工传动有限公司",
        "91330206MA2GQ8T71C",
        "专精特新小巨人",
        "第五批",
        "2023-07-14",
        "2026-07-13",
        "浙江省",
        "工业和信息化部",
      ],
      [
        "绍兴新材科技股份有限公司",
        "91330600MA29C6K82X",
        "省级专精特新",
        "2025年度",
        "2025-04-18",
        "2028-04-17",
        "浙江省",
        "浙江省经信厅",
      ],
    ],
  },
  "EXT-R005": {
    summary:
      "按企业查询诉讼、执行、失信等司法事件，支持授信调查和存量客户风险核验。",
    purposes: ["司法核验", "授信调查", "风险预警"],
    products: ["企业司法风险监测", "授信尽调外部信息核验"],
    params: [
      {
        name: "enterprise_name",
        cn: "企业名称",
        value: "北辰精密制造有限公司",
      },
      { name: "event_type", cn: "司法事件类型", value: "全部" },
    ],
    fields: [
      [
        "enterprise_name",
        "企业名称",
        "涉案企业的登记名称",
        "string",
        "北辰精密制造有限公司",
      ],
      [
        "case_no",
        "案号",
        "人民法院编制的案件编号",
        "string",
        "（2026）京0105执3281号",
      ],
      [
        "case_type",
        "案件类型",
        "民事诉讼、被执行、失信等事件类型",
        "string",
        "被执行",
      ],
      [
        "case_role",
        "诉讼地位",
        "企业在案件中的原告、被告或被执行人身份",
        "string",
        "被执行人",
      ],
      [
        "court_name",
        "受理法院",
        "受理或执行案件的法院名称",
        "string",
        "北京市朝阳区人民法院",
      ],
      [
        "case_amount",
        "涉案金额（元）",
        "公开文书披露的案件金额",
        "decimal",
        "760000",
      ],
      ["filing_date", "立案日期", "法院登记立案日期", "date", "2026-07-21"],
      [
        "risk_level",
        "风险等级",
        "依据事件类型和金额计算的风险等级",
        "string",
        "中",
      ],
    ],
    rows: [
      [
        "北辰精密制造有限公司",
        "（2026）京0105执3281号",
        "被执行",
        "被执行人",
        "北京市朝阳区人民法院",
        760000,
        "2026-07-21",
        "中",
      ],
      [
        "华东数控系统股份有限公司",
        "（2026）沪0115民初8226号",
        "民事诉讼",
        "被告",
        "上海市浦东新区人民法院",
        125000,
        "2026-07-17",
        "低",
      ],
      [
        "南方新材科技有限公司",
        "（2026）粤0305执912号",
        "被执行",
        "被执行人",
        "深圳市南山区人民法院",
        2680000,
        "2026-07-09",
        "高",
      ],
    ],
  },
  "EXT-R006": {
    summary:
      "提供经营异常名录、严重违法失信和信用修复状态，用于企业准入排查与持续风险监测。",
    purposes: ["经营风险排查", "失信核验", "准入筛选"],
    products: ["企业司法风险监测", "专精特新企业营销名单"],
    params: [
      {
        name: "credit_code",
        cn: "统一社会信用代码",
        value: "91440300MA5F8Q2L7R",
      },
      { name: "status", cn: "当前状态", value: "未移出" },
    ],
    fields: [
      [
        "enterprise_name",
        "企业名称",
        "被列入异常或失信名录的企业名称",
        "string",
        "南方新材科技有限公司",
      ],
      [
        "credit_code",
        "统一社会信用代码",
        "企业法定身份识别代码",
        "string",
        "91440300MA5F8Q2L7R",
      ],
      [
        "abnormal_type",
        "异常类型",
        "列入经营异常或严重失信的具体类型",
        "string",
        "未按期公示年度报告",
      ],
      [
        "listed_date",
        "列入日期",
        "主管部门作出列入决定的日期",
        "date",
        "2026-06-30",
      ],
      [
        "decision_authority",
        "决定机关",
        "作出列入决定的市场监管部门",
        "string",
        "深圳市市场监督管理局",
      ],
      [
        "current_status",
        "当前状态",
        "仍在名录、已移出或信用修复中",
        "string",
        "仍在名录",
      ],
      ["removed_date", "移出日期", "完成信用修复并移出名录的日期", "date", "—"],
      [
        "severity",
        "严重程度",
        "按异常类型和持续时间划分的程度",
        "string",
        "中",
      ],
    ],
    rows: [
      [
        "南方新材科技有限公司",
        "91440300MA5F8Q2L7R",
        "未按期公示年度报告",
        "2026-06-30",
        "深圳市市场监督管理局",
        "仍在名录",
        "—",
        "中",
      ],
      [
        "华北机电设备有限公司",
        "91120116MA06Q4D91T",
        "通过登记住所无法联系",
        "2026-05-18",
        "天津市滨海新区市场监管局",
        "信用修复中",
        "—",
        "中",
      ],
      [
        "苏州精密传感科技有限公司",
        "91320594MA1Y8R6H2Q",
        "公示信息隐瞒真实情况",
        "2025-11-06",
        "苏州工业园区市场监管局",
        "已移出",
        "2026-07-02",
        "低",
      ],
    ],
  },
  "EXT-R007": {
    summary:
      "实时聚合新闻、媒体与公开网络信息，识别企业主体、事件类型和情绪倾向，用于舆情监测与风险提示。",
    purposes: ["舆情监测", "事件识别", "风险提示"],
    products: ["舆情风险预警", "核心企业供应链风险画像"],
    params: [
      { name: "keyword", cn: "企业或主题关键词", value: "新能源装备" },
      { name: "sentiment", cn: "情感倾向", value: "负面" },
    ],
    fields: [
      [
        "article_id",
        "资讯编号",
        "聚合平台生成的资讯唯一标识",
        "string",
        "NEWS-20260803-0931",
      ],
      [
        "title",
        "资讯标题",
        "公开报道或舆情信息标题",
        "string",
        "某装备企业回应项目延期传闻",
      ],
      [
        "subject_name",
        "关联主体",
        "从正文识别并关联的企业或机构",
        "string",
        "华北新能源装备有限公司",
      ],
      [
        "publish_time",
        "发布时间",
        "原始信息首次公开的时间",
        "datetime",
        "2026-08-03 08:42:16",
      ],
      [
        "source_name",
        "来源媒体",
        "发布该信息的媒体或网站名称",
        "string",
        "新华财经",
      ],
      [
        "event_type",
        "事件类型",
        "经营、司法、安全、治理等事件标签",
        "string",
        "项目延期",
      ],
      [
        "sentiment",
        "情感倾向",
        "模型识别的正面、中性或负面倾向",
        "string",
        "负面",
      ],
      [
        "risk_score",
        "舆情风险分",
        "综合来源、传播和事件性质计算的分值",
        "integer",
        "82",
      ],
      [
        "summary",
        "内容摘要",
        "对原文核心事实的脱敏摘要",
        "string",
        "企业称将调整交付计划并评估影响",
      ],
    ],
    rows: [
      [
        "NEWS-20260803-0931",
        "某装备企业回应项目延期传闻",
        "华北新能源装备有限公司",
        "2026-08-03 08:42:16",
        "新华财经",
        "项目延期",
        "负面",
        82,
        "企业称将调整交付计划并评估影响",
      ],
      [
        "NEWS-20260803-0718",
        "精密制造企业获新一轮融资",
        "南京芯联装备股份有限公司",
        "2026-08-03 07:18:40",
        "经济参考报",
        "融资进展",
        "正面",
        18,
        "企业完成融资并拟扩建生产基地",
      ],
      [
        "NEWS-20260802-1645",
        "监管通报产品质量抽检结果",
        "南方新材科技有限公司",
        "2026-08-02 16:45:09",
        "地方市场监管部门",
        "质量监管",
        "负面",
        76,
        "部分批次产品被要求限期整改",
      ],
    ],
  },
};

function ResourceManagement({ appMode = false }: { appMode?: boolean }) {
  const standardized =
    typeof window !== "undefined" &&
    localStorage.getItem("icbc-standard-resource") === "published";
  const standardResource: any = {
    id: "RES-JUD-2026-081",
    name: "标准企业司法风险数据服务",
    category: "企业数据 / 风险信息",
    type: "标准API",
    supplier: "法信数据（经标准化）",
    frequency: "每日",
    quality: "A",
    auth: "已授权",
    fields:
      "enterprise_id、risk_event_type、risk_level、event_date、case_amount、source_record_id",
    standardized: true,
  };
  const resourceRows = standardized
    ? [standardResource, ...resources]
    : resources;
  const [selected, setSelected] = useState<any>(resourceRows[0]);
  const [open, setOpen] = useState(false);
  const [apiRan, setApiRan] = useState(false);
  const profile = firstPageResourceProfiles[selected.id];
  const fallbackFields = String(selected.fields || "")
    .split(/、|，|,/)
    .filter(Boolean)
    .map((cn: string, i: number) => [
      `field_${i + 1}`,
      cn,
      `用于描述${cn}的业务属性`,
      "string",
      "演示值",
    ]);
  const fieldRows = (profile?.fields || fallbackFields).map(
    (x: any, i: number) => ({
      key: i,
      field: x[0],
      cn: x[1],
      meaning: x[2],
      type: x[3],
      example: x[4],
    }),
  );
  const previewRows = (profile?.rows || []).map((values: any[], i: number) =>
    Object.fromEntries([
      ["key", i + 1],
      ...fieldRows.map((f: any, j: number) => [f.field, values[j]]),
    ]),
  );
  const previewColumns = fieldRows.map((f: any) => ({
    title: (
      <span>
        {f.cn}
        <small>{f.field}</small>
      </span>
    ),
    dataIndex: f.field,
    render: (value: any) =>
      typeof value === "number" ? value.toLocaleString() : String(value ?? "—"),
  }));
  const followResource = () => {
    const old = JSON.parse(
      localStorage.getItem("icbc-follow-resources") || "[]",
    );
    if (!old.includes(selected.id))
      localStorage.setItem(
        "icbc-follow-resources",
        JSON.stringify([...old, selected.id]),
      );
    message.success(`已收藏资源“${selected.name}”`);
  };
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            {appMode ? "外数资源" : "资源目录管理"} <DemoTag />
          </h1>
          <p>
            {appMode
              ? "按业务对象和用途查找可使用的外部数据资源"
              : "目录治理服务于产品查重、建设编排、字段映射和资源血缘"}
          </p>
        </div>
        {!appMode && (
          <Button type="primary" icon={<PlusOutlined />}>
            新增目录节点
          </Button>
        )}
      </div>
      <div className="resource-page">
        <Card className="tree-card" title="外数资源目录">
          <Input prefix={<SearchOutlined />} placeholder="搜索目录节点" />
          <Tree
            treeData={directoryTree}
            defaultExpandedKeys={["enterprise", "basic", "registry", "risk"]}
            selectedKeys={["profile"]}
          />
        </Card>
        <Card className="resource-list">
          <div className="table-tools">
            <Space>
              <Select
                defaultValue="全部资源类型"
                options={[
                  "全部资源类型",
                  "数据库",
                  "API接口",
                  "标准API",
                  "数据文件",
                  "查询服务",
                  "研究报告",
                  "聚合资讯",
                ].map((value) => ({ value }))}
              />
              <Select
                defaultValue="全部授权状态"
                options={["全部授权状态", "已授权", "限域授权", "采购中"].map(
                  (value) => ({ value }),
                )}
              />
            </Space>
            <span>已筛选 {resourceRows.length} 项资源</span>
          </div>
          <Table
            size="small"
            pagination={{ pageSize: 8, showSizeChanger: false }}
            dataSource={resourceRows}
            rowKey="id"
            onRow={(r) => ({
              onClick: () => {
                setSelected(r);
                setApiRan(false);
                setOpen(true);
              },
            })}
            columns={[
              {
                title: "资源名称",
                dataIndex: "name",
                render: (x: string, r: any) => (
                  <div>
                    <a>{x}</a>
                    {r.standardized && (
                      <Tag color="success" className="standard-tag">
                        已标准化·可编排
                      </Tag>
                    )}
                    <small className="cell-sub">
                      {r.id} · {r.category}
                    </small>
                  </div>
                ),
              },
              {
                title: "类型",
                dataIndex: "type",
                render: (x: string) => <Tag>{x}</Tag>,
              },
              { title: "供应商", dataIndex: "supplier" },
              { title: "更新频率", dataIndex: "frequency" },
              {
                title: "质量",
                dataIndex: "quality",
                render: (x: string) => <Tag color="green">{x}级</Tag>,
              },
              {
                title: "授权",
                dataIndex: "auth",
                render: (x: string) => <StatusTag status={x} />,
              },
            ]}
          />
        </Card>
      </div>
      <Drawer
        width={920}
        open={open}
        onClose={() => setOpen(false)}
        title="外数资源详情与在线预览"
        extra={
          appMode ? (
            <Button icon={<StarOutlined />} onClick={followResource}>
              加入自选
            </Button>
          ) : null
        }
      >
        <div className="resource-detail-hero">
          <DatabaseOutlined />
          <div>
            <Tag>{selected.type}</Tag>
            <h2>{selected.name}</h2>
            <p>{selected.id} · 最近更新 2026-07-31 06:00</p>
          </div>
        </div>
        <Descriptions
          bordered
          size="small"
          column={3}
          items={[
            { label: "资源供应商", children: selected.supplier },
            {
              label: "获取方式",
              children: selected.type === "数据文件" ? "文件获取" : "接口调用",
            },
            { label: "更新频率", children: selected.frequency },
            { label: "数据质量", children: selected.quality + "级" },
            { label: "覆盖地域", children: "全国" },
            {
              label: "授权状态",
              children: <StatusTag status={selected.auth} />,
            },
          ]}
        />
        <Tabs
          className="resource-preview-tabs"
          items={[
            {
              key: "intro",
              label: "业务说明",
              children: (
                <>
                  <h3 className="drawer-title">这项资源包含什么</h3>
                  <p>
                    {profile?.summary ||
                      `${selected.fields}等业务字段，可通过受控资源服务获取。`}
                  </p>
                  <h3 className="drawer-title">适合解决什么问题</h3>
                  <Space wrap>
                    {(profile?.purposes || ["资源检索", "业务核验"]).map(
                      (x: string) => (
                        <Tag key={x}>{x}</Tag>
                      ),
                    )}
                  </Space>
                  <h3 className="drawer-title">已支撑产品</h3>
                  <List
                    dataSource={
                      profile?.products ||
                      initialProducts.slice(0, 2).map((p) => p.name)
                    }
                    renderItem={(name: string) => (
                      <List.Item>
                        <a>{name}</a>
                        <Tag>正式产品</Tag>
                      </List.Item>
                    )}
                  />
                </>
              ),
            },
            {
              key: "fields",
              label: `字段字典（${fieldRows.length}）`,
              children: (
                <>
                  <div className="field-help">
                    <SafetyCertificateOutlined />
                    <span>
                      <b>字段名称同时展示中文业务含义</b>
                      ，业务人员无需阅读供应商接口文档即可理解每一列数据。
                    </span>
                  </div>
                  <Table
                    size="small"
                    pagination={false}
                    dataSource={fieldRows}
                    columns={[
                      {
                        title: "标准字段名",
                        dataIndex: "field",
                        render: (x: string) => <code>{x}</code>,
                      },
                      {
                        title: "中文名称",
                        dataIndex: "cn",
                        render: (x: string) => <b>{x}</b>,
                      },
                      { title: "字段说明", dataIndex: "meaning" },
                      { title: "数据类型", dataIndex: "type" },
                      { title: "样例值", dataIndex: "example" },
                    ]}
                  />
                </>
              ),
            },
            {
              key: "sample",
              label: "数据样例预览",
              children: (
                <>
                  <p className="preview-note">
                    展示与当前资源口径一致的脱敏演示数据；正式使用时仅返回当前岗位获授权的字段与地域范围。
                  </p>
                  {previewRows.length ? (
                    <Table
                      size="small"
                      scroll={{ x: Math.max(900, fieldRows.length * 155) }}
                      pagination={false}
                      dataSource={previewRows}
                      columns={previewColumns}
                    />
                  ) : (
                    <Empty description="该资源未配置首屏演示样例" />
                  )}
                </>
              ),
            },
            {
              key: "api",
              label: "在线查询与接口试用",
              children: profile ? (
                <div className="api-console">
                  <Card
                    size="small"
                    title={`${selected.type}模拟请求`}
                    extra={<Tag color="warning">脱敏沙箱</Tag>}
                  >
                    <Form layout="vertical">
                      <Row gutter={12}>
                        {profile.params.map((p: any) => (
                          <Col span={12} key={p.name}>
                            <Form.Item label={`${p.cn} ${p.name}`}>
                              <Input defaultValue={p.value} />
                            </Form.Item>
                          </Col>
                        ))}
                      </Row>
                      <Button
                        type="primary"
                        icon={<ApiOutlined />}
                        onClick={() => setApiRan(true)}
                      >
                        执行在线查询
                      </Button>
                    </Form>
                  </Card>
                  {apiRan ? (
                    <Card size="small" title="200 OK · 当前资源响应">
                      <pre>
                        {JSON.stringify(
                          {
                            code: "0000",
                            message: "success",
                            resource_id: selected.id,
                            data: previewRows.slice(0, 2),
                          },
                          null,
                          2,
                        )}
                      </pre>
                    </Card>
                  ) : (
                    <Empty description="填写当前资源的查询条件后，在线查看对应结果" />
                  )}
                </div>
              ) : (
                <Empty description="该资源不在首屏演示范围，仅提供目录说明" />
              ),
            },
          ]}
        />
      </Drawer>
    </div>
  );
}

type IntakeRow = Record<string, unknown>;

const resourceFieldAliases: Record<string, string[]> = {
  enterprise_name: ["enterprise_name", "company_name", "企业名称", "企业全称", "主体名称", "公司名称"],
  enterprise_id: ["enterprise_id", "credit_code", "统一社会信用代码", "统一信用代码", "企业id", "主体id"],
  registered_capital: ["registered_capital", "reg_capital", "注册资本", "注册资金"],
  establish_date: ["establish_date", "成立日期", "注册日期", "设立日期"],
  legal_representative: ["legal_representative", "法人", "法定代表人"],
  industry_code: ["industry_code", "行业代码", "所属行业", "行业"],
  region_code: ["region_code", "地区", "省份", "注册地", "行政区划"],
  bid_count: ["bid_count", "招投标次数", "中标次数", "投标次数"],
  bid_amount: ["bid_amount", "中标金额", "招投标金额", "项目金额"],
  bid_detail: ["bid_detail", "项目名称", "招投标明细", "中标项目"],
  qualification_level: ["qualification_level", "专精特新资质", "企业资质", "资质等级"],
  provincial_qualification: ["provincial_qualification", "省级专精特新", "省市级资质"],
  risk_event_type: ["risk_event_type", "case_type", "司法事件类型", "案件类型", "风险事件"],
  risk_level: ["risk_level", "risk_grade", "风险等级", "风险级别"],
  event_date: ["event_date", "filing_date", "事件日期", "立案日期", "发生日期"],
  case_amount: ["case_amount", "amount_wan", "涉案金额", "执行金额", "案件金额"],
  environmental_penalty: ["environmental_penalty", "环保处罚", "环境处罚"],
  abnormal_operation: ["abnormal_operation", "经营异常", "经营异常标识"],
};

const standardFieldNames: Record<string, string> = {
  enterprise_name: "企业名称", enterprise_id: "统一社会信用代码", registered_capital: "注册资本",
  establish_date: "成立日期", legal_representative: "法定代表人", industry_code: "行业代码",
  region_code: "行政区划", bid_count: "招投标次数", bid_amount: "中标金额", bid_detail: "招投标项目明细",
  qualification_level: "国家级专精特新资质", provincial_qualification: "省市级专精特新资质",
  risk_event_type: "司法风险事件类型", risk_level: "风险等级", event_date: "事件日期",
  case_amount: "涉案金额", environmental_penalty: "环保处罚", abnormal_operation: "经营异常",
};

const stockResourceProfiles = [
  { id: "RES-ENT-001", name: "全国企业工商登记数据库", topic: "企业基本信息", fields: ["enterprise_name","enterprise_id","registered_capital","establish_date","legal_representative","industry_code","region_code"], coverage: "全国企业法人", history: "2010年至今", frequency: "每日", quality: 96, auth: "全行经营管理与风险用途", cost: "存量合同" },
  { id: "RES-BID-006", name: "全国招投标项目明细库", topic: "企业经营行为", fields: ["enterprise_name","enterprise_id","bid_count","bid_amount","bid_detail","event_date"], coverage: "全国公开招投标", history: "近5年", frequency: "每日", quality: 91, auth: "全行营销与研究用途", cost: "存量合同" },
  { id: "RES-QUAL-011", name: "国家级专精特新企业名单", topic: "企业资质", fields: ["enterprise_name","enterprise_id","qualification_level","industry_code","region_code"], coverage: "国家级名单", history: "历年批次", frequency: "按批次", quality: 99, auth: "全行内部使用", cost: "公共数据" },
  { id: "RES-JUD-021", name: "企业司法涉诉与执行信息", topic: "企业风险信息", fields: ["enterprise_name","enterprise_id","risk_event_type","risk_level","event_date","case_amount"], coverage: "全国企业司法事项", history: "近8年", frequency: "准实时", quality: 94, auth: "授信、风险及合规用途", cost: "存量合同" },
  { id: "RES-OPS-027", name: "企业经营异常信息", topic: "企业风险信息", fields: ["enterprise_name","enterprise_id","abnormal_operation","event_date"], coverage: "全国市场主体", history: "2015年至今", frequency: "每日", quality: 97, auth: "全行内部使用", cost: "公共数据" },
  { id: "RES-ENV-034", name: "企业环保处罚信息库", topic: "企业风险信息", fields: ["enterprise_name","enterprise_id","environmental_penalty","event_date","case_amount"], coverage: "重点省市，约82%企业", history: "近3年", frequency: "每周", quality: 84, auth: "风险监测用途", cost: "试用资源" },
];

function normalizeFieldName(value: string) {
  const clean = String(value || "").trim().toLowerCase().replace(/[\s_\-（）()]/g, "");
  const hit = Object.entries(resourceFieldAliases).find(([, aliases]) =>
    aliases.some((x) => x.toLowerCase().replace(/[\s_\-（）()]/g, "") === clean),
  );
  return hit?.[0] || `unmapped:${value}`;
}

function ResourceOnboarding({ go }: { go: (v: View) => void }) {
  const [mode, setMode] = useState<"excel" | "api">("excel");
  const [rows, setRows] = useState<IntakeRow[]>([]);
  const [fileName, setFileName] = useState("");
  const [apiOpen, setApiOpen] = useState(false);
  const [apiUrl, setApiUrl] = useState("demo://supplier/enterprise-resource-v1");
  const [apiJson, setApiJson] = useState("");
  const [loading, setLoading] = useState(false);
  const [meta, setMeta] = useState({ topic: "企业综合经营与风险信息", coverage: "全国企业", history: "近3年", frequency: "每日", quality: 88, auth: "营销、授信核验与风险监测" });
  const [active, setActive] = useState("intake");

  const sourceFields = useMemo(() => rows.length ? Object.keys(rows[0]) : [], [rows]);
  const mappedFields = useMemo(() => sourceFields.map((source) => ({
    source, standard: normalizeFieldName(source), sample: rows[0]?.[source],
  })), [sourceFields, rows]);
  const canonicalFields = useMemo(() => mappedFields.filter(x => !x.standard.startsWith("unmapped:")).map(x => x.standard), [mappedFields]);
  const unmappedFields = mappedFields.filter(x => x.standard.startsWith("unmapped:"));
  const comparisons = useMemo(() => stockResourceProfiles.map((r) => {
    const overlap = canonicalFields.filter(f => r.fields.includes(f));
    const added = canonicalFields.filter(f => !r.fields.includes(f));
    const missing = r.fields.filter(f => !canonicalFields.includes(f));
    const fieldScore = canonicalFields.length ? Math.round(overlap.length / canonicalFields.length * 100) : 0;
    const semanticScore = meta.topic.includes(r.topic.replace("企业", "")) || r.topic.includes("企业") ? Math.min(98, fieldScore + 12) : fieldScore;
    const coverageScore = meta.coverage.includes("全国") && r.coverage.includes("全国") ? 100 : meta.coverage === r.coverage ? 100 : 62;
    const qualityDelta = meta.quality - r.quality;
    const frequencyScore = meta.frequency === r.frequency ? 100 : 65;
    // 内容重叠率是查重主指标。主题、覆盖、质量和时效只能帮助定位候选资源，
    // 不能在没有任何重复字段时“制造”重复率。
    const supportingScore = Math.round(semanticScore * .35 + coverageScore * .3 + Math.min(100, r.quality) * .2 + frequencyScore * .15);
    const score = overlap.length ? fieldScore : 0;
    return { ...r, overlap, added, missing, fieldScore, semanticScore, coverageScore, frequencyScore, supportingScore, qualityDelta, score };
  }).sort((a,b) => b.score - a.score || b.supportingScore - a.supportingScore), [canonicalFields, meta]);
  const coveredSet = new Set(comparisons.flatMap(x => x.overlap));
  const repeated = canonicalFields.filter(x => coveredSet.has(x));
  const newFields = canonicalFields.filter(x => !coveredSet.has(x));
  const qualityUpgrades = comparisons.filter(x => x.overlap.length && x.qualityDelta >= 5);
  const best = comparisons[0];

  const loadRows = (data: IntakeRow[], name: string, nextMode: "excel" | "api") => {
    if (!data.length || !Object.keys(data[0] || {}).length) return message.error("未识别到可比较的结构化记录");
    setRows(data.slice(0, 200)); setFileName(name); setMode(nextMode); setActive("profile");
    message.success(`已读取 ${data.length} 条记录、${Object.keys(data[0]).length} 个字段，开始库内查重`);
  };
  const parseFile = async (file: File) => {
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json<IntakeRow>(sheet, { defval: "" });
      loadRows(data, file.name, "excel");
    } catch { message.error("文件解析失败，请检查Excel首行是否为字段名称"); }
    return false;
  };
  const demoRows = [
    { 企业名称: "华北新能源装备有限公司", 统一社会信用代码: "91110108DEMO000001", 注册资本: 28000, 成立日期: "2023-05-16", 招投标次数: 12, 中标金额: 18600, 项目名称: "动力电池产线扩建项目", 省级专精特新: "北京市专精特新中小企业", 环保处罚: "无" },
    { 企业名称: "南京芯联装备股份有限公司", 统一社会信用代码: "91320100DEMO000002", 注册资本: 16500, 成立日期: "2021-09-08", 招投标次数: 8, 中标金额: 9200, 项目名称: "半导体设备采购项目", 省级专精特新: "江苏省专精特新中小企业", 环保处罚: "2026-05整改完成" },
  ];
  const readApi = async () => {
    setLoading(true);
    try {
      let payload: any;
      if (apiJson.trim()) payload = JSON.parse(apiJson);
      else if (apiUrl.startsWith("demo://")) payload = { data: demoRows };
      else {
        const response = await fetch(apiUrl, { method: "GET", headers: { Accept: "application/json" } });
        if (!response.ok) throw new Error(String(response.status));
        payload = await response.json();
      }
      const data = Array.isArray(payload) ? payload : payload.data || payload.results || payload.items || [payload];
      loadRows(data, apiUrl, "api"); setApiOpen(false);
    } catch { message.error("API读取失败：请检查地址、跨域许可或粘贴脱敏JSON响应后重试"); }
    finally { setLoading(false); }
  };

  const conclusion = !rows.length ? "等待接入资源"
    : canonicalFields.length === 0
    ? `暂无法自动判断是否重复：${unmappedFields.length} 个字段尚未映射到行内标准语义，需人工定义后重新查重`
    : repeated.length === 0
    ? `未发现直接重复：${canonicalFields.length} 个已识别标准字段均未被存量资源覆盖，建议进入新增资源价值、质量、授权与成本评估`
    : newFields.length
    ? `部分重叠：建议复用 ${repeated.length} 项已有字段，仅评估引入 ${newFields.length} 项实质新增字段`
    : qualityUpgrades.length ? "内容高度重复但质量存在增量：建议开展同口径样本测试，评估替换存量资源" : "核心内容已被存量资源覆盖：原则上复用现有资源，不重复引入";

  return <div className="resource-dedupe-page">
    <div className="page-heading"><div><h1>外数资源引入查重 <DemoTag /></h1><p>在采购与接入前真实读取样本，完成目录定位、字段映射、存量比对和差异引入判断</p></div><Space><Tag color="red">前置采购关口</Tag><Button onClick={() => { setRows([]); setFileName(""); setActive("intake"); }}>新建查重任务</Button></Space></div>
    <Card className="dedupe-principle"><FileSearchOutlined /><div><b>查重对象是拟引入外数资源，而非产品名称</b><span>系统比较业务语义、主体覆盖、标准字段、时间范围、质量时效、授权与成本，输出可执行的差异清单。</span></div><div className="dedupe-chain"><span>读取</span><em>→</em><span>映射</span><em>→</em><span>比对</span><em>→</em><span>决策</span></div></Card>
    <Tabs activeKey={active} onChange={setActive} items={[
      { key:"intake", label:"1 资源接入", children:<Row gutter={16}>
        <Col span={15}><Card title="选择真实读取方式" className="intake-method-card"><Radio.Group value={mode} onChange={e=>setMode(e.target.value)} optionType="button" buttonStyle="solid" options={[{label:"Excel / CSV 文件",value:"excel"},{label:"API 接口",value:"api"}]} />
          {mode === "excel" ? <Upload.Dragger accept=".xlsx,.xls,.csv" multiple={false} showUploadList={false} beforeUpload={parseFile}><p className="ant-upload-drag-icon"><CloudUploadOutlined /></p><p className="ant-upload-text">拖入或选择一份真实Excel/CSV样本</p><p className="ant-upload-hint">首行为字段名；文件在浏览器内解析，本原型最多读取前200条用于查重</p></Upload.Dragger> : <div className="api-intake"><ApiOutlined /><h3>读取供应商或内部代理API</h3><p>演示接口可直接运行；真实接口需允许跨域访问。生产环境应由行内服务端代理保管鉴权信息。</p><Button type="primary" onClick={()=>setApiOpen(true)}>配置并读取API</Button></div>}
          <Button block icon={<ExperimentOutlined />} onClick={()=>loadRows(demoRows,"企业综合数据包_查重演示.xlsx","excel")}>使用内置虚拟数据真实运行查重</Button>
        </Card></Col>
        <Col span={9}><Card title="本次比较所需元数据"><Form layout="vertical"><Form.Item label="拟引入资源主题"><Input value={meta.topic} onChange={e=>setMeta({...meta,topic:e.target.value})}/></Form.Item><Form.Item label="主体与地域覆盖"><Input value={meta.coverage} onChange={e=>setMeta({...meta,coverage:e.target.value})}/></Form.Item><Row gutter={8}><Col span={12}><Form.Item label="历史范围"><Input value={meta.history} onChange={e=>setMeta({...meta,history:e.target.value})}/></Form.Item></Col><Col span={12}><Form.Item label="更新频率"><Select value={meta.frequency} onChange={frequency=>setMeta({...meta,frequency})} options={["准实时","每日","每周","每月","按批次"].map(value=>({value}))}/></Form.Item></Col></Row><Form.Item label="样本质量评分"><InputNumber min={0} max={100} value={meta.quality} onChange={quality=>setMeta({...meta,quality:Number(quality||0)})} style={{width:"100%"}}/></Form.Item><Form.Item label="拟授权用途"><Input value={meta.auth} onChange={e=>setMeta({...meta,auth:e.target.value})}/></Form.Item></Form></Card></Col>
      </Row>},
      { key:"profile", disabled:!rows.length, label:"2 读取与字段映射", children:<><Card className="read-result" title="真实读取结果" extra={<Space><Tag color="success">已读取</Tag><span>{fileName}</span></Space>}><div className="metric-strip"><div><b>{rows.length}</b><span>样本记录</span></div><div><b>{sourceFields.length}</b><span>原始字段</span></div><div><b>{canonicalFields.length}</b><span>映射成功</span></div><div className="warn"><b>{unmappedFields.length}</b><span>待人工定义</span></div></div><Table size="small" pagination={false} scroll={{x:Math.max(900,sourceFields.length*150)}} dataSource={rows.slice(0,5).map((r,i)=>({...r,key:i}))} columns={sourceFields.map(f=>({title:f,dataIndex:f,ellipsis:true}))}/></Card>
        <Card title="供应商字段 → 行内标准字段" extra={<Tag>别名词典＋样例结构匹配</Tag>}><Table size="small" pagination={false} dataSource={mappedFields.map((x,i)=>({...x,key:i}))} columns={[{title:"原始字段",dataIndex:"source",render:x=><code>{x}</code>},{title:"样例值",dataIndex:"sample",ellipsis:true},{title:"行内标准字段",dataIndex:"standard",render:(x:string)=><Tag color={x.startsWith("unmapped:")?"orange":"green"}>{x.startsWith("unmapped:")?"待人工定义":`${standardFieldNames[x]}（${x}）`}</Tag>},{title:"匹配结论",render:(_:any,r:any)=>r.standard.startsWith("unmapped:")?<span>进入新增字段评估，不自动判为新增资源</span>:<span>可用于跨供应商内容查重</span>}]} /></Card><div className="pane-actions"><Button type="primary" onClick={()=>setActive("compare")}>运行存量资源比对</Button></div></>},
      { key:"compare", disabled:!rows.length, label:"3 存量比较", children:<><div className="compare-hero"><div><Tag color="processing">实时计算</Tag><h2>系统已将 {canonicalFields.length} 个标准字段与 {stockResourceProfiles.length} 项存量资源逐项比较</h2><p>{canonicalFields.length ? "圆环仅表示与最相关存量资源的内容重叠率；主题、覆盖、质量和时效作为辅助依据单独展示。" : "当前没有字段映射到行内标准语义，暂不计算内容重复率；请先完成人工字段定义。"}</p></div><Progress type="circle" percent={best?.score||0} size={88} format={p=>`${p}%`} /></div><Table rowKey="id" pagination={false} dataSource={comparisons} expandable={{expandedRowRender:(r:any)=><div className="score-grid"><div><span>内容重叠</span><Progress percent={r.fieldScore} size="small" /></div><div><span>业务主题</span><Progress percent={r.semanticScore} size="small" /></div><div><span>覆盖范围</span><Progress percent={r.coverageScore} size="small" /></div><div><span>更新时效</span><Progress percent={r.frequencyScore} size="small" /></div><div><span>拟引入字段</span><b>{canonicalFields.length} 项</b></div><div><span>实际重复</span><b>{r.overlap.length} 项：{r.overlap.length?r.overlap.map((f:string)=>standardFieldNames[f]).join("、"):"无"}</b></div><div><span>仅拟引入方有</span><b>{r.added.length} 项</b></div><div><span>质量评分差</span><b>{r.qualityDelta>0?"+":""}{r.qualityDelta}（拟引入−存量）</b></div></div>}} columns={[{title:"存量资源",render:(_:any,r:any)=><div><b>{r.name}</b><small className="cell-sub">{r.id} · {r.topic}</small></div>},{title:"内容重叠率",dataIndex:"score",render:(x:number,r:any)=><div><Progress percent={x} size="small"/><small className="cell-sub">{r.overlap.length}/{canonicalFields.length || 0} 个已识别字段</small></div>},{title:"重复标准字段",render:(_:any,r:any)=><span><b>{r.overlap.length}</b> 项{r.overlap.length?`：${r.overlap.slice(0,3).map((f:string)=>standardFieldNames[f]).join("、")}${r.overlap.length>3?"等":""}`:"：无"}</span>},{title:"辅助相关性",render:(_:any,r:any)=><span>{r.supportingScore}%<br/><small>仅用于候选定位，不计入重复率</small></span>},{title:"质量评分差",render:(_:any,r:any)=><Tag color={r.qualityDelta>=5?"green":r.qualityDelta<=-5?"orange":"default"}>{r.qualityDelta>0?"+":""}{r.qualityDelta}（拟引入−存量）</Tag>},{title:"初步判断",render:(_:any,r:any)=>r.overlap.length===0?<Tag>无直接重叠</Tag>:r.fieldScore>=80?<Tag color="red">高度重叠</Tag>:<Tag color="orange">部分重叠</Tag>}]} /><div className="review-rule">展开任一存量资源，可查看内容重叠、主题、覆盖、时效及质量差的逐项依据。</div><div className="pane-actions"><Button type="primary" onClick={()=>setActive("decision")}>生成差异清单与引入建议</Button></div></>},
      { key:"decision", disabled:!rows.length, label:"4 查重结论", children:<><div className="decision-banner"><AuditOutlined/><div><Tag color={canonicalFields.length===0?"warning":repeated.length===0?"blue":newFields.length?"orange":"green"}>查重结论</Tag><h2>{conclusion}</h2><p>结论只由实际字段覆盖关系触发；主题、覆盖、质量和时效不会单独形成“重复”判断。</p></div></div><Row gutter={14}><Col span={6}><Card className="decision-stat"><Statistic title="识别标准字段" value={canonicalFields.length}/><span>可进入跨资源比较</span></Card></Col><Col span={6}><Card className="decision-stat"><Statistic title="已有资源覆盖" value={repeated.length}/><span>{repeated.length?"优先复用，不重复采购":"未发现可直接复用字段"}</span></Card></Col><Col span={6}><Card className="decision-stat"><Statistic title="实质新增字段" value={newFields.length}/><span>{canonicalFields.length?"进入价值与合规评估":"待完成标准语义映射"}</span></Card></Col><Col span={6}><Card className="decision-stat"><Statistic title="待定义字段" value={unmappedFields.length}/><span>人工定义后重新计算</span></Card></Col></Row>
        <Row gutter={14}><Col span={14}><Card title="逐字段差异清单"><Table size="small" pagination={false} dataSource={mappedFields.map((x,i)=>({...x,key:i}))} columns={[{title:"拟引入字段",dataIndex:"source"},{title:"标准内容",render:(_:any,r:any)=>r.standard.startsWith("unmapped:")?"尚未建立标准语义":standardFieldNames[r.standard]},{title:"存量覆盖",render:(_:any,r:any)=>{const hits=comparisons.filter(c=>c.fields.includes(r.standard));return hits.length?hits.slice(0,2).map(h=>h.name).join("；"):"当前库内未发现"}},{title:"处置建议",render:(_:any,r:any)=>r.standard.startsWith("unmapped:")?<Tag color="orange">人工定义后复查</Tag>:coveredSet.has(r.standard)?<Tag color="green">复用存量</Tag>:<Tag color="red">评估差异引入</Tag>}]} /></Card></Col><Col span={10}><Card title="采购与接入建议"><Timeline items={[{color:"green",children:<><b>复用已有内容</b><p>{repeated.length?repeated.map(f=>standardFieldNames[f]).join("、"):"暂无"}</p></>},{color:"orange",children:<><b>仅评估增量内容</b><p>{newFields.length?newFields.map(f=>standardFieldNames[f]).join("、"):"未识别到实质新增字段"}</p></>},{color:qualityUpgrades.length?"blue":"gray",children:<><b>质量替换测试</b><p>{qualityUpgrades.length?`新资源质量高于 ${qualityUpgrades.map(x=>x.name).join("、")}，建议抽取同主体同日期样本测试准确性与时效性。`:"当前样本未形成明确质量替代优势"}</p></>},{color:"red",children:<><b>授权与成本人工核验</b><p>核对是否允许模型加工、衍生产品、全行下发及拆分报价；系统不以字段重合替代合同判断。</p></>}]} /><Button type="primary" block onClick={()=>message.success("查重报告已保存，进入资源引入人工评审")}>提交资源引入评审</Button><Button block style={{marginTop:8}} onClick={()=>go("manage-resources")}>查看存量资源目录</Button></Card></Col></Row></>},
    ]}/>
    <Modal open={apiOpen} onCancel={()=>setApiOpen(false)} onOk={readApi} confirmLoading={loading} okText="读取并开始查重" title="配置API数据源" width={760}><Form layout="vertical"><Form.Item label="接口地址"><Input value={apiUrl} onChange={e=>setApiUrl(e.target.value)} prefix={<ApiOutlined/>}/></Form.Item><div className="api-security-note"><SafetyCertificateOutlined/><span><b>演示边界：</b>可读取允许浏览器跨域访问的JSON接口；用户名、密码和令牌不得填入页面。正式环境应由行内服务端代理、密钥中心和接口白名单完成鉴权。</span></div><Form.Item label="或粘贴脱敏JSON响应（支持 data / results / items 数组）"><Input.TextArea rows={9} value={apiJson} onChange={e=>setApiJson(e.target.value)} placeholder={'{"data":[{"企业名称":"示例企业","环保处罚":"无"}]}'}/></Form.Item></Form></Modal>
  </div>;
}

function ResourceStandardizationLegacy({ go }: { go: (v: View) => void }) {
  const saved =
    typeof window !== "undefined" &&
    localStorage.getItem("icbc-standard-resource") === "published";
  const [current, setCurrent] = useState(saved ? 5 : 0);
  const [uploaded, setUploaded] = useState(saved);
  const [published, setPublished] = useState(saved);
  const [fileName, setFileName] = useState(
    saved ? "企业司法风险数据_202608.xlsx" : "",
  );
  const [mapping, setMapping] = useState<any>({
    company_name: "enterprise_name",
    credit_code: "enterprise_id",
    case_type: "risk_event_type",
    risk_grade: "risk_level",
    filing_date: "event_date",
    amount_wan: "case_amount",
  });
  const [mappingStatus, setMappingStatus] = useState<Record<string, string>>({
    company_name: "待复核",
    credit_code: "已确认",
    case_type: "待复核",
    risk_grade: "口径冲突",
    filing_date: "待复核",
    amount_wan: "待复核",
  });
  const [mappingDetail, setMappingDetail] = useState<any>(null);
  const [newFieldOpen, setNewFieldOpen] = useState(false);
  const sourceRows = [
    {
      company_name: "北京智造科技有限公司",
      credit_code: "91110108MA01X8KQ2N",
      case_type: "被执行",
      risk_grade: "高",
      filing_date: "2026/07/18",
      amount_wan: "126.50",
    },
    {
      company_name: "京华新能源装备股份有限公司",
      credit_code: "91110114MA0207TY6C",
      case_type: "裁判文书",
      risk_grade: "中",
      filing_date: "2026-07-12",
      amount_wan: "38",
    },
    {
      company_name: "北辰精密制造有限公司",
      credit_code: "",
      case_type: "经营异常",
      risk_grade: "低",
      filing_date: "2026.07.09",
      amount_wan: "0",
    },
  ];
  const stepNames = [
    "原始资源接入",
    "结构与元数据剖析",
    "标准字段映射",
    "主体与代码标准化",
    "质量与授权校验",
    "标准服务发布",
  ];
  const fieldMeta: any = {
    company_name: {
      cn: "企业名称",
      definition: "工商登记或司法文书中的企业全称",
      type: "string",
      unit: "—",
      confidence: 96,
      name: 98,
      semantic: 96,
      object: 100,
      structure: 92,
      sampleScore: 94,
      history: "3个同类资源已采用",
      transform: "文本清洗",
    },
    credit_code: {
      cn: "统一社会信用代码",
      definition: "企业唯一社会信用标识",
      type: "string(18)",
      unit: "—",
      confidence: 99,
      name: 100,
      semantic: 99,
      object: 100,
      structure: 100,
      sampleScore: 98,
      history: "法信数据历史映射 V2.1",
      transform: "格式校验",
    },
    case_type: {
      cn: "案件类型",
      definition: "供应商司法事项分类",
      type: "string",
      unit: "代码集",
      confidence: 88,
      name: 91,
      semantic: 86,
      object: 100,
      structure: 82,
      sampleScore: 80,
      history: "2个司法资源已采用",
      transform: "枚举映射",
    },
    risk_grade: {
      cn: "风险等级",
      definition: "供应商依据自有模型形成的高、中、低等级",
      type: "string",
      unit: "供应商口径",
      confidence: 91,
      name: 96,
      semantic: 78,
      object: 100,
      structure: 94,
      sampleScore: 85,
      history: "曾否决直接映射1次",
      transform: "需核验计算口径",
      conflict: "名称相似，但供应商等级含义与行内风险等级规则不同",
    },
    filing_date: {
      cn: "立案日期",
      definition: "司法机关登记立案的日期",
      type: "date",
      unit: "YYYY/MM/DD",
      confidence: 93,
      name: 94,
      semantic: 95,
      object: 100,
      structure: 88,
      sampleScore: 90,
      history: "5个资源已采用",
      transform: "统一为 YYYY-MM-DD",
    },
    amount_wan: {
      cn: "涉案金额",
      definition: "案件涉及的标的金额，供应商以万元交付",
      type: "decimal(18,2)",
      unit: "万元",
      confidence: 86,
      name: 82,
      semantic: 91,
      object: 100,
      structure: 89,
      sampleScore: 83,
      history: "同供应商接口已采用",
      transform: "万元 × 10,000 → 元",
    },
  };
  const standardOptions = [
    "enterprise_name",
    "enterprise_id",
    "risk_event_type",
    "risk_level",
    "event_date",
    "case_amount",
    "supplier_risk_grade",
  ];
  const upload = (name = "企业司法风险数据_202608.xlsx") => {
    setUploaded(true);
    setFileName(name);
    setCurrent(1);
    message.success("文件上传成功，已进入隔离接入区");
  };
  const next = () => setCurrent((v) => Math.min(5, v + 1));
  const reset = () => {
    localStorage.removeItem("icbc-standard-resource");
    setUploaded(false);
    setPublished(false);
    setFileName("");
    setCurrent(0);
    setMappingStatus({
      company_name: "待复核",
      credit_code: "已确认",
      case_type: "待复核",
      risk_grade: "口径冲突",
      filing_date: "待复核",
      amount_wan: "待复核",
    });
    message.success("已重置本次演示任务");
  };
  const publish = () => {
    localStorage.setItem("icbc-standard-resource", "published");
    setPublished(true);
    message.success("标准资源服务已入库上架");
  };
  const footer = (label = "保存并进入下一步") => (
    <div className="onboarding-actions">
      <span>
        <CheckCircleOutlined /> 当前步骤操作将完整记录处理日志和版本
      </span>
      <Space>
        {current > 0 && !published && (
          <Button onClick={() => setCurrent((v) => v - 1)}>上一步</Button>
        )}
        <Button
          type="primary"
          onClick={current === 5 ? publish : next}
          disabled={!uploaded}
        >
          {current === 5 ? "确认入库并上架" : label}
        </Button>
      </Space>
    </div>
  );
  const panels = [
    <div className="upload-stage" key="upload">
      <Card className="upload-zone">
        <CloudUploadOutlined />
        <h2>上传外数资源样本</h2>
        <p>
          支持
          Excel、CSV、API文档及结构化数据包。原始文件进入隔离接入区，不直接供产品调用。
        </p>
        <Space>
          <Upload
            accept=".xlsx,.xls,.csv"
            showUploadList={false}
            beforeUpload={(f) => {
              upload(f.name);
              return false;
            }}
          >
            <Button type="primary" icon={<CloudUploadOutlined />}>
              选择本地Excel
            </Button>
          </Upload>
          <Button onClick={() => upload()} icon={<ExperimentOutlined />}>
            使用演示文件
          </Button>
        </Space>
      </Card>
      <Card title="准入说明" className="intake-note">
        <Timeline
          items={[
            { children: "采购或授权验收已完成" },
            { children: "保留供应商原始交付与批次信息" },
            { children: "完成标准化前仅限接入人员查看" },
            { children: "通过质量、授权校验后方可发布" },
          ]}
        />
      </Card>
    </div>,
    <div key="profile">
      <Row gutter={14}>
        <Col span={16}>
          <Card
            title="自动识别的文件结构"
            extra={<Tag color="success">剖析完成 · 3条样本</Tag>}
          >
            <Table
              size="small"
              pagination={false}
              scroll={{ x: 900 }}
              dataSource={sourceRows.map((x, i) => ({ ...x, key: i }))}
              columns={Object.keys(sourceRows[0]).map((k) => ({
                title: k,
                dataIndex: k,
                render: (x: string) => <code>{x || "NULL"}</code>,
              }))}
            />
          </Card>
        </Col>
        <Col span={8}>
          <Card title="资源元数据">
            <Descriptions
              column={1}
              size="small"
              bordered
              items={[
                { label: "文件名称", children: fileName },
                { label: "资源对象", children: "企业司法风险事件" },
                { label: "交付方式", children: "Excel批量文件" },
                { label: "记录规模", children: "12,680条（演示）" },
                { label: "识别字段", children: "6个" },
                { label: "更新周期", children: "每日增量" },
                { label: "原始区位置", children: "RAW/JUD/20260803" },
              ]}
            />
          </Card>
        </Col>
      </Row>
      {footer()}
    </div>,
    <div key="mapping" className="mapping-workspace">
      <div className="match-summary">
        <div>
          <b>6</b>
          <span>解析字段</span>
        </div>
        <div>
          <b>6</b>
          <span>系统推荐</span>
        </div>
        <div>
          <b>
            {Object.values(mappingStatus).filter((x) => x === "已确认").length}
          </b>
          <span>已人工确认</span>
        </div>
        <div className="warn">
          <b>
            {Object.values(mappingStatus).filter((x) => x === "待复核").length}
          </b>
          <span>待人工复核</span>
        </div>
        <div className="danger">
          <b>
            {
              Object.values(mappingStatus).filter((x) => x === "口径冲突")
                .length
            }
          </b>
          <span>口径冲突</span>
        </div>
      </div>
      <Card className="match-method" size="small">
        <RobotOutlined />
        <div>
          <b>多层匹配引擎已完成候选推荐</b>
          <span>
            字段标识与别名 → 业务语义 → 对象与数据结构 → 样例值验证 →
            历史映射复用
          </span>
        </div>
        <Button
          onClick={() =>
            message.success("已重新运行匹配：推荐6项，发现口径冲突1项")
          }
          icon={<ReloadOutlined />}
        >
          重新自动匹配
        </Button>
      </Card>
      <Card
        title="供应商字段与行内标准字段复核"
        extra={
          <span className="review-rule">系统只推荐，人工确认后方可生效</span>
        }
      >
        <Table
          size="small"
          scroll={{ x: 1120 }}
          pagination={false}
          dataSource={Object.entries(mapping).map(([source, target], i) => ({
            key: source,
            source,
            sample: Object.values(sourceRows[0])[i],
            target,
            ...fieldMeta[source],
          }))}
          columns={[
            {
              title: "供应商原始字段",
              width: 175,
              render: (_: any, r: any) => (
                <div className="source-field">
                  <code>{r.source}</code>
                  <small>
                    {r.cn} · {r.type}
                  </small>
                </div>
              ),
            },
            {
              title: "样例值",
              dataIndex: "sample",
              width: 170,
              ellipsis: true,
            },
            {
              title: "系统推荐的标准字段",
              width: 230,
              render: (_: any, r: any) => (
                <Select
                  value={r.target}
                  style={{ width: 210 }}
                  onChange={(v) => {
                    setMapping((m: any) => ({ ...m, [r.source]: v }));
                    setMappingStatus((s) => ({ ...s, [r.source]: "待复核" }));
                  }}
                  options={standardOptions.map((value) => ({ value }))}
                />
              ),
            },
            {
              title: "标准化处理",
              dataIndex: "transform",
              width: 160,
              render: (x: string) => <Tag>{x}</Tag>,
            },
            {
              title: "综合置信度",
              dataIndex: "confidence",
              width: 135,
              render: (x: number) => (
                <Progress
                  percent={x}
                  size="small"
                  strokeColor={x >= 90 ? "#198754" : "#d28b00"}
                />
              ),
            },
            {
              title: "复核状态",
              width: 105,
              render: (_: any, r: any) => (
                <StatusTag status={mappingStatus[r.source]} />
              ),
            },
            {
              title: "操作",
              fixed: "right" as const,
              width: 215,
              render: (_: any, r: any) => (
                <Space size={4}>
                  <Button size="small" onClick={() => setMappingDetail(r)}>
                    查看依据
                  </Button>
                  <Button
                    size="small"
                    type={
                      mappingStatus[r.source] === "已确认"
                        ? "default"
                        : "primary"
                    }
                    disabled={mappingStatus[r.source] === "口径冲突"}
                    onClick={() => {
                      setMappingStatus((s) => ({ ...s, [r.source]: "已确认" }));
                      message.success(`${r.cn}映射已人工确认`);
                    }}
                  >
                    确认
                  </Button>
                  <Dropdown
                    menu={{
                      items: [
                        { key: "new", label: "发起新字段评估" },
                        { key: "extension", label: "保留为供应商扩展字段" },
                      ],
                      onClick: ({ key }) => {
                        if (key === "new") setNewFieldOpen(true);
                        else {
                          setMapping((m: any) => ({
                            ...m,
                            [r.source]: "supplier_risk_grade",
                          }));
                          setMappingStatus((s) => ({
                            ...s,
                            [r.source]: "已确认",
                          }));
                          message.success("已登记为供应商扩展字段");
                        }
                      },
                    }}
                  >
                    <Button size="small">其他</Button>
                  </Dropdown>
                </Space>
              ),
            },
          ]}
        />
      </Card>
      <div className="mapping-bottom">
        <Card size="small" title="复核后的标准化样例">
          <div className="sample-compare">
            <div>
              <small>原始值</small>
              <code>amount_wan = 126.50</code>
            </div>
            <em>× 10,000</em>
            <div>
              <small>标准值</small>
              <code>case_amount = 1,265,000.00 CNY</code>
            </div>
          </div>
        </Card>
        <Card size="small" title="进入下一步的门禁">
          <p>
            所有字段必须完成确认、扩展字段登记或新增字段审批；
            <b>口径冲突未处置时禁止进入下一步。</b>
          </p>
        </Card>
      </div>
      <Modal
        open={!!mappingDetail}
        onCancel={() => setMappingDetail(null)}
        footer={
          <Space>
            <Button onClick={() => setMappingDetail(null)}>关闭</Button>
            {mappingDetail && (
              <Button
                type="primary"
                disabled={mappingStatus[mappingDetail.source] === "口径冲突"}
                onClick={() => {
                  setMappingStatus((s) => ({
                    ...s,
                    [mappingDetail.source]: "已确认",
                  }));
                  setMappingDetail(null);
                  message.success("映射已确认并记录复核人、时间与依据");
                }}
              >
                确认该映射
              </Button>
            )}
          </Space>
        }
        title="自动匹配依据与人工复核"
      >
        {mappingDetail && (
          <div className="match-evidence">
            <div className="evidence-head">
              <div>
                <small>供应商字段</small>
                <b>
                  {mappingDetail.source} · {mappingDetail.cn}
                </b>
              </div>
              <em>推荐</em>
              <div>
                <small>行内标准字段</small>
                <b>{mappingDetail.target}</b>
              </div>
            </div>
            {mappingDetail.conflict && (
              <div className="conflict-note">
                <SafetyCertificateOutlined />
                <span>
                  <b>检测到口径冲突</b>
                  {mappingDetail.conflict}
                </span>
              </div>
            )}
            <Descriptions
              bordered
              column={2}
              size="small"
              items={[
                {
                  label: "业务定义",
                  children: mappingDetail.definition,
                  span: 2,
                },
                { label: "数据类型", children: mappingDetail.type },
                { label: "单位/代码集", children: mappingDetail.unit },
                { label: "历史映射", children: mappingDetail.history, span: 2 },
              ]}
            />
            <h4>综合置信度构成</h4>
            <div className="score-grid">
              {[
                ["名称与别名", mappingDetail.name],
                ["语义相似度", mappingDetail.semantic],
                ["业务对象", mappingDetail.object],
                ["类型与单位", mappingDetail.structure],
                ["样例值验证", mappingDetail.sampleScore],
              ].map((x) => (
                <div key={String(x[0])}>
                  <span>{x[0]}</span>
                  <Progress percent={Number(x[1])} size="small" />
                </div>
              ))}
            </div>
            <div className="audit-note">
              复核结论将记录复核人、时间、供应商版本、候选差异和影响范围。
            </div>
          </div>
        )}
      </Modal>
      <Modal
        open={newFieldOpen}
        onCancel={() => setNewFieldOpen(false)}
        onOk={() => {
          setNewFieldOpen(false);
          message.success("新字段评估申请已提交，进入标准治理审批");
        }}
        title="发起新字段评估"
        okText="提交治理审批"
      >
        <Form layout="vertical">
          <Form.Item label="拟新增标准字段名称" required>
            <Input defaultValue="企业司法风险供应商评级" />
          </Form.Item>
          <Row gutter={12}>
            <Col span={12}>
              <Form.Item label="建议字段层级">
                <Select
                  defaultValue="供应商扩展字段"
                  options={[
                    "企业级标准字段",
                    "领域标准字段",
                    "供应商扩展字段",
                    "产品派生指标",
                  ].map((value) => ({ value }))}
                />
              </Form.Item>
            </Col>
            <Col span={12}>
              <Form.Item label="所属业务对象">
                <Input defaultValue="企业司法风险事件" />
              </Form.Item>
            </Col>
          </Row>
          <Form.Item label="新增理由">
            <Input.TextArea defaultValue="现有标准库无同口径字段；需先保留供应商来源口径，待跨资源复用验证后再评估升级。" />
          </Form.Item>
        </Form>
      </Modal>
      <div className="onboarding-actions">
        <span>
          <SafetyCertificateOutlined />{" "}
          所有自动推荐均需人工复核，口径冲突须先处置
        </span>
        <Space>
          <Button onClick={() => setCurrent((v) => v - 1)}>上一步</Button>
          <Button
            onClick={() => {
              setMappingStatus(
                Object.fromEntries(
                  Object.keys(mappingStatus).map((k) => [
                    k,
                    k === "risk_grade" ? "口径冲突" : "已确认",
                  ]),
                ),
              );
              message.success("高置信度且无冲突的字段已批量确认");
            }}
          >
            批量确认无冲突项
          </Button>
          <Button
            type="primary"
            disabled={Object.values(mappingStatus).some((x) => x !== "已确认")}
            onClick={next}
          >
            完成复核并进入下一步
          </Button>
        </Space>
      </div>
    </div>,
    <div key="entity">
      <Row gutter={14}>
        <Col span={15}>
          <Card title="主体关联与代码统一">
            <div className="standard-flow">
              <div>
                <b>原始企业名称 / 信用代码</b>
                <small>供应商字段</small>
              </div>
              <em>→</em>
              <div>
                <b>企业主体关联</b>
                <small>名称清洗＋信用代码校验</small>
              </div>
              <em>→</em>
              <div className="active">
                <b>标准企业主体ID</b>
                <small>enterprise_id</small>
              </div>
            </div>
            <Table
              pagination={false}
              size="small"
              dataSource={sourceRows.map((x, i) => ({
                ...x,
                key: i,
                match: i === 2 ? "名称模糊匹配" : "信用代码精确匹配",
                id: [
                  "ENT-110108-003821",
                  "ENT-110114-002176",
                  "ENT-110105-008923",
                ][i],
              }))}
              columns={[
                { title: "原始企业名称", dataIndex: "company_name" },
                {
                  title: "原始信用代码",
                  dataIndex: "credit_code",
                  render: (x: string) => x || <Tag color="warning">缺失</Tag>,
                },
                { title: "关联方式", dataIndex: "match" },
                {
                  title: "标准主体ID",
                  dataIndex: "id",
                  render: (x: string) => <code>{x}</code>,
                },
                {
                  title: "结果",
                  render: () => <Tag color="success">已关联</Tag>,
                },
              ]}
            />
          </Card>
        </Col>
        <Col span={9}>
          <Card title="代码与口径转换">
            <List
              dataSource={[
                "司法事件类型：供应商代码 → 行内5类事件",
                "风险等级：高/中/低 → 3/2/1",
                "日期：多格式 → YYYY-MM-DD",
                "金额：万元 → 人民币元",
                "行政区域：文本 → GB/T 2260代码",
              ]}
              renderItem={(x, i) => (
                <List.Item>
                  <CheckCircleFilled className="pass-icon" />
                  <span>{x}</span>
                  <Tag color={i < 4 ? "success" : "default"}>
                    {i < 4 ? "已转换" : "无需转换"}
                  </Tag>
                </List.Item>
              )}
            />
          </Card>
        </Col>
      </Row>
      {footer()}
    </div>,
    <div key="quality">
      <Row gutter={14}>
        <Col span={15}>
          <Card title="数据质量校验">
            <div className="quality-grid">
              {[
                ["完整性", "98.7%", 99],
                ["格式合规率", "100%", 100],
                ["主体匹配率", "99.4%", 99],
                ["重复记录率", "0.3%", 97],
                ["时效达标率", "100%", 100],
                ["口径一致性", "96.8%", 97],
              ].map((x) => (
                <div key={String(x[0])}>
                  <Progress type="circle" percent={Number(x[2])} size={72} />
                  <b>{x[0]}</b>
                  <small>{x[1]}</small>
                </div>
              ))}
            </div>
            <div className="quality-result">
              <CheckCircleFilled />
              <div>
                <b>资源质量等级：A级，达到正式发布标准</b>
                <p>
                  发现164条主体标识缺失记录，已通过企业名称模糊匹配补全158条；剩余6条进入异常数据队列。
                </p>
              </div>
            </div>
          </Card>
        </Col>
        <Col span={9}>
          <Card title="授权与使用边界">
            <Descriptions
              column={1}
              bordered
              size="small"
              items={[
                { label: "授权地域", children: "境内机构" },
                { label: "使用目的", children: "营销筛选、授信核验、风险监测" },
                { label: "禁止用途", children: "对外转售、个人用途" },
                { label: "授权期限", children: "2026-08-01 至 2027-07-31" },
                { label: "数据留存", children: "原始数据12个月，结果24个月" },
                {
                  label: "敏感等级",
                  children: <Tag color="orange">内部数据</Tag>,
                },
                { label: "校验结论", children: <StatusTag status="通过" /> },
              ]}
            />
          </Card>
        </Col>
      </Row>
      {footer()}
    </div>,
    <div key="publish">
      <Row gutter={14}>
        <Col span={16}>
          <Card title="标准资源服务发布清单">
            <Descriptions
              bordered
              column={2}
              items={[
                { label: "服务名称", children: "标准企业司法风险数据服务" },
                { label: "资源编号", children: "RES-JUD-2026-081" },
                { label: "标准对象", children: "企业司法风险事件" },
                { label: "服务形态", children: "标准API＋标准数据表" },
                {
                  label: "输入契约",
                  children: <code>enterprise_id, start_date, end_date</code>,
                  span: 2,
                },
                {
                  label: "输出契约",
                  children: (
                    <code>
                      risk_event_type, risk_level, event_date, case_amount
                    </code>
                  ),
                  span: 2,
                },
                {
                  label: "关联适配器",
                  children: "法信司法风险Excel适配器 V1.0",
                },
                { label: "质量等级", children: <Tag color="success">A级</Tag> },
                {
                  label: "可用能力",
                  children: "司法风险核验、风险事件统计、名单排除",
                },
                { label: "可用规则", children: "近N月重大司法风险排除规则" },
              ]}
            />
            <Card
              size="small"
              className="service-contract"
              title="标准服务调用示例"
            >
              <code>
                GET
                /standard-resources/v1/enterprise-risk?enterprise_id=ENT-110108-003821&amp;months=3
              </code>
            </Card>
          </Card>
        </Col>
        <Col span={8}>
          <Card title="发布影响">
            <Timeline
              items={[
                { color: "green", children: "生成标准资源目录项" },
                { color: "green", children: "登记原始—标准字段血缘" },
                { color: "green", children: "绑定质量与授权策略" },
                { color: "green", children: "开放给能力模块调用" },
                { color: "gray", children: "产品建设中心可检索并编排" },
              ]}
            />
            {published ? (
              <Result
                status="success"
                title="已入库上架"
                subTitle="该资源已成为可编排的标准资源服务"
                extra={
                  <Space direction="vertical">
                    <Button
                      type="primary"
                      onClick={() => go("manage-resources")}
                    >
                      前往资源目录查看
                    </Button>
                    <Button onClick={() => go("build")}>
                      进入产品建设中心
                    </Button>
                    <Button type="link" onClick={reset}>
                      重新演示
                    </Button>
                  </Space>
                }
              />
            ) : (
              footer()
            )}
          </Card>
        </Col>
      </Row>
    </div>,
  ];
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            资源接入与标准化 <DemoTag />
          </h1>
          <p>
            将供应商异构数据转换为具有统一语义、标准契约和受控权限的可编排资源服务
          </p>
        </div>
        <Space>
          {uploaded && (
            <Tag color="processing">接入任务：ONB-2026-0803-018</Tag>
          )}
          <Button onClick={reset}>重置演示</Button>
        </Space>
      </div>
      <Card className="standardization-principle">
        <DatabaseOutlined />
        <div>
          <b>原始资源保留，标准服务发布</b>
          <span>
            供应商物理结构不必完全相同；通过字段映射、资源适配器和标准数据契约，向上层能力模块提供统一输入。
          </span>
        </div>
        <div className="layer-chain">
          <span>原始资源</span>
          <em>→</em>
          <span>标准化层</span>
          <em>→</em>
          <span>能力与规则</span>
          <em>→</em>
          <span>产品组装</span>
        </div>
      </Card>
      <Steps
        className="onboarding-steps"
        current={current}
        items={stepNames.map((title, i) => ({
          title,
          description:
            i < current ? "已完成" : i === current ? "当前步骤" : "待处理",
        }))}
      />
      <div className="onboarding-panel">{panels[current]}</div>
    </div>
  );
}

function Assets({
  deposited,
  go,
}: {
  deposited: boolean;
  go: (v: View) => void;
}) {
  const [type, setType] = useState("product");
  const [detail, setDetail] = useState<any>(null);
  const [builderOpen, setBuilderOpen] = useState(false);
  const [builderStep, setBuilderStep] = useState(0);
  const [builderKind, setBuilderKind] = useState("module");
  const [builderName, setBuilderName] = useState("产业事件影响评分");
  const [builderTested, setBuilderTested] = useState(false);
  const [manualAssets, setManualAssets] = useState<any[]>([]);
  useEffect(() => {
    try {
      const savedDraft = localStorage.getItem("icbc-manual-asset-draft");
      if (savedDraft) setManualAssets([JSON.parse(savedDraft)]);
    } catch {}
  }, []);
  const templates = [
    {
      id: "T001",
      name: "专精特新企业营销母版",
      type: "产品母版",
      source: "专精特新企业营销名单",
      input: "企业主体＋资质＋行为＋风险",
      output: "营销名单＋行动提示",
      calls: 18,
      version: "V3.2",
      owner: "总行数据管理部",
    },
    {
      id: "T002",
      name: "企业风险持续监测母版",
      type: "产品母版",
      source: "企业司法风险监测",
      input: "客户名单＋风险事件",
      output: "分级预警＋核验清单",
      calls: 12,
      version: "V2.6",
      owner: "总行风险管理部",
    },
    {
      id: "T003",
      name: "区域产业客户画像母版",
      type: "产品母版",
      source: "区域产业链客户画像",
      input: "区域＋产业链＋企业信息",
      output: "客户画像＋产业图谱",
      calls: 9,
      version: "V2.1",
      owner: "总行数据管理部",
    },
    {
      id: "T004",
      name: "企业动态机会与风险评估母版",
      type: "产品母版",
      source: "全国产业链机会与风险识别",
      input: "产业事件＋企业基本面＋信用信息＋客户经营内数",
      output: "三维指数＋变化归因＋综合判断＋行动任务",
      calls: 7,
      version: "V3.1",
      owner: "总行数据管理部",
    },
    {
      id: "T005",
      name: "授信内外部信息核验母版",
      type: "产品母版",
      source: "授信尽调外部信息核验",
      input: "授信申请＋客户主数据＋多源外部事实",
      output: "核验结论＋冲突项＋复核建议",
      calls: 11,
      version: "V2.4",
      owner: "总行授信管理部",
    },
    {
      id: "T006",
      name: "客户经营趋势预测母版",
      type: "产品母版",
      source: "企业经营活跃度监测",
      input: "结算时序＋授信变化＋经营信号",
      output: "趋势预测＋需求概率＋行动窗口",
      calls: 5,
      version: "V1.8",
      owner: "总行公司金融业务部",
    },
  ];
  const standardProducts = [
    { id: "P001", name: "专精特新企业营销名单", type: "标准产品", source: "浙江分行成果标准化", input: "地区＋行业＋企业资质＋风险信息", output: "营销名单＋入选依据＋行动提示", calls: 36, version: "V3.2", owner: "总行公司金融业务部", status: "正常运行" },
    { id: "P002", name: "企业司法风险监测", type: "标准产品", source: "多分行风险产品归并", input: "客户名单＋司法风险事件", output: "分级预警＋核验清单", calls: 28, version: "V2.6", owner: "总行风险管理部", status: "正常运行" },
    { id: "P003", name: "产业客户机会识别", type: "标准产品", source: "四川分行产业客户产品", input: "地区＋行业＋期限＋风险阈值", output: "候选名单＋证据链＋行动提示", calls: 14, version: "V1.0", owner: "总行数据管理部", status: "试点运行" },
  ];
  const taskPacks = [
    {
      id: "TP001",
      name: "产业事件研判与业务分派",
      type: "任务能力包",
      source: "全国产业链机会与风险识别",
      input: "行业/地区＋企业主体＋观察期限",
      output: "企业名单＋行业简报＋营销/风险任务",
      calls: 23,
      version: "V1.3",
      owner: "总行数据管理部",
      trigger: "识别产业事件影响并形成业务行动",
      assembly: "4项能力＋4条规则＋3类成果模板",
    },
    {
      id: "TP002",
      name: "授信尽调信息核验",
      type: "任务能力包",
      source: "授信尽调外部信息核验",
      input: "授信申请＋企业主体＋核验范围",
      output: "事实核验表＋冲突项＋复核任务",
      calls: 31,
      version: "V2.1",
      owner: "总行授信管理部",
      trigger: "对授信客户开展多源事实交叉核验",
      assembly: "5项能力＋6条规则＋2类成果模板",
    },
    {
      id: "TP003",
      name: "专精特新客户发现与触达",
      type: "任务能力包",
      source: "专精特新企业营销名单",
      input: "地区＋行业＋资质口径＋排除条件",
      output: "营销清单＋客户画像＋触达话术",
      calls: 46,
      version: "V2.4",
      owner: "总行公司金融业务部",
      trigger: "筛选专精特新潜客并分派客户经理",
      assembly: "6项能力＋5条规则＋3类成果模板",
    },
  ];
  const ruleData = [
    ...rules.map((x) => ({
      ...x,
      type: "通用规则",
      source: "全行规则库",
      calls: Math.max(6, 42 - Number(x.id.slice(1)) * 3),
    })),
    ...(deposited
      ? [
          {
            id: "R-ZJ-001",
            name: "浙江地方专精特新资质识别规则",
            input: "地方资质名单",
            output: "标准资质标签",
            calls: 0,
            version: "V1.0",
            owner: "浙江分行",
            type: "通用规则",
            source: "浙江分行专精特新企业筛选产品",
          },
        ]
      : []),
  ];
  const moduleData = modules.map((x) => ({
    ...x,
    type: "能力模块",
    source: "全行标准产品",
  }));
  const data =
    type === "product" ? standardProducts : type === "template" ? [...manualAssets.filter(x=>x.type === "产品母版"), ...templates] : type === "module" ? [...manualAssets.filter(x=>x.type === "能力模块"), ...moduleData] : ruleData;
  const labels: any = {
    product: [
      "标准产品",
      "经功能、适配和运行验证后正式发布的完整业务成果，可由业务人员直接配置参数并生成名单、指标、风险提示或行动建议。",
    ],
    template: [
      "产品母版",
      "功能：固化成熟产品的完整业务链路、输入输出和组件组合。使用：在建设路径中整体引用，再调整地区、行业、期限、阈值及地方数据映射，快速生成新产品。",
    ],
    module: [
      "能力模块",
      "功能：封装主体关联、指标计算、名单筛选、司法核验等独立处理能力。使用：在流程编排画布中按输入输出契约直接拖入，并配置参数、版本和异常处理。",
    ],
    rule: [
      "通用规则",
      "功能：统一管理可复用的判断条件、计算口径和参数边界。使用：在规则配置器中选择规则，填写地区、期限、阈值等参数，经测试后随产品版本发布。",
    ],
  };
  const detailInfo = (x: any) =>
    x.type === "标准产品"
      ? {
          purpose: `围绕“${x.name}”这一明确业务任务，调用已验证的资源、模块和规则，直接形成可供业务人员使用的结果。`,
          scene: "客户营销、授信支持、风险监测等具体业务场景",
          usage: "业务应用门户 → 外数产品 → 配置地区、行业、期限和阈值 → 生成结果",
          boundary: "核心流程和关键风险口径锁定；仅开放经审批的业务参数，所有调用留痕并受权限控制",
          params: "适用地区、目标行业、观察期限、筛选阈值、结果范围",
          relations: `由“${x.source}”转化形成，累计调用 ${x.calls} 次`,
        }
      : x.type === "任务能力包"
      ? {
          purpose: `面向“${x.trigger}”这一完整任务，统一调用所需能力、规则与交付模板，减少重复拆解和逐项选配。`,
          scene: "自然语言任务已有成熟标准路径，需要快速形成可审查的产品编排草案",
          usage: "产品建设中心 → 产品流程画布 → 任务能力包 → 一键应用 → 逐节点人工确认",
          boundary: "仅生成建设草案，不自动发布；锁定口径、风险排除和低置信度结果必须人工复核",
          params: "任务对象、地区、行业、观察期限、阈值、交付岗位与成果格式",
          relations: `${x.assembly}；来源于“${x.source}”，累计调用 ${x.calls} 次`,
        }
      : x.type === "产品母版"
      ? {
          purpose: `以“${x.source}”的成熟建设成果为基础，复用完整业务任务、处理链路与结果交付方式。`,
          scene: "同类业务任务在新地区、新行业或新机构快速复制",
          usage:
            "产品建设中心 → 建设路径 → 选择“基于产品母版适配” → 配置地域参数和地方数据映射",
          boundary:
            "允许调整地区、行业、期限、阈值和地方数据来源；核心业务口径、主流程与审批版本不可直接修改",
          params: "适用地域、行业范围、观察期限、筛选阈值、地方名单映射",
          relations: `来源于“${x.source}”，当前被 ${x.calls} 个正式或候选产品引用`,
        }
      : x.type === "能力模块"
        ? {
            purpose: `将“${x.name}”封装为可独立调用的标准处理能力，对上游输入执行确定性加工并返回标准结果。`,
            scene: "不同产品重复出现相同处理节点时直接复用",
            usage:
              "产品建设中心 → 能力匹配 → 在流程编排画布中加入模块 → 校验输入输出契约",
            boundary:
              "可配置版本、输入字段映射、超时和异常处理；模块内部标准算法由维护部门统一管理",
            params: "输入字段、调用版本、批量规模、异常处理方式、结果返回字段",
            relations: `由“${x.source}”沉淀，当前被 ${x.calls} 个产品调用`,
          }
        : {
            purpose: `统一执行“${x.name}”对应的业务判断条件，确保不同产品使用一致口径。`,
            scene: "名单筛选、风险排除、资质判断及阈值控制",
            usage:
              "产品建设中心 → 规则配置 → 选择规则 → 填写参数 → 运行样本测试 → 随产品版本发布",
            boundary:
              "仅允许在批准范围内调整地区、期限和阈值；规则定义、适用对象与关键口径需走变更审批",
            params: "适用地域、观察期、比较符、阈值、空值处理、规则优先级",
            relations: `来自“${x.source}”，当前关联 ${x.calls} 个产品`,
          };
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            能力资产中心 <DemoTag />
          </h1>
          <p>按照复用粒度统一管理标准产品、产品母版、能力模块和通用规则</p>
        </div>
        <Button
          type="primary"
          icon={<BuildOutlined />}
          onClick={() => {
            setBuilderStep(0);
            setBuilderTested(false);
            setBuilderOpen(true);
          }}
        >
          手工建设能力资产
        </Button>
      </div>
      <div className="asset-level-map">
        <div><span>01</span><small>完整成果</small><b>标准产品</b><p>验证后直接服务业务</p></div>
        <em>拆解 / 复用</em>
        <div><span>02</span><small>完整骨架</small><b>产品母版</b><p>调整参数快速复制</p></div>
        <em>结构化提取</em>
        <div><span>03</span><small>独立环节</small><b>能力模块</b><p>标准输入输出，可组合调用</p></div>
        <em>条件参数化</em>
        <div><span>04</span><small>判断口径</small><b>通用规则</b><p>地区、行业、期限与阈值</p></div>
      </div>
      <Row gutter={14} className="asset-summary">
        <Col span={6}>
          <MetricCard
            title="标准产品"
            value={standardProducts.length}
            trend="经验证可直接使用"
            onClick={() => setType("product")}
          />
        </Col>
        <Col span={6}>
          <MetricCard
            title="产品母版"
            value={templates.length}
            trend="复用完整业务骨架"
            onClick={() => setType("template")}
          />
        </Col>
        <Col span={6}>
          <MetricCard
            title="能力模块"
            value={modules.length}
            trend="独立加工、按契约调用"
            onClick={() => setType("module")}
          />
        </Col>
        <Col span={6}>
          <MetricCard
            title="通用规则"
            value={rules.length + (deposited ? 1 : 0)}
            trend={deposited ? "刚刚回沉 1 项" : "条件参数化复用"}
            onClick={() => setType("rule")}
          />
        </Col>
      </Row>
      <Card>
        <Tabs
          activeKey={type}
          onChange={setType}
          items={[
            { key: "product", label: `标准产品（${standardProducts.length}）` },
            { key: "template", label: `产品母版（${templates.length}）` },
            { key: "module", label: `能力模块（${modules.length}）` },
            { key: "rule", label: `通用规则（${rules.length + (deposited ? 1 : 0)}）` },
          ]}
        />
        <div className="asset-type-note">
          <b>{labels[type][0]}</b>
          <span>{labels[type][1]}</span>
        </div>
        <Table
          dataSource={data}
          rowKey="id"
          pagination={{ pageSize: 8 }}
          columns={[
            {
              title: "能力资产",
              dataIndex: "name",
              render: (x: string, r: any) => (
                <div>
                  <a onClick={() => setDetail(r)}>{x}</a>
                  <small className="cell-sub">
                    {r.id} · {r.type}
                  </small>
                </div>
              ),
            },
            { title: "来源产品", dataIndex: "source" },
            { title: "输入", dataIndex: "input" },
            { title: "输出", dataIndex: "output" },
            {
              title: "调用产品",
              dataIndex: "calls",
              render: (x: number) => x + " 个",
            },
            { title: "当前版本", dataIndex: "version" },
            { title: "维护主体", dataIndex: "owner" },
            { title: "状态", render: (_:any,r:any) => <StatusTag status={r.status || "正常运行"} /> },
            {
              title: "操作",
              render: (_: any, r: any) => (
                <Button size="small" onClick={() => setDetail(r)}>
                  查看详情
                </Button>
              ),
            },
          ]}
        />
      </Card>
      <Drawer
        width={760}
        open={!!detail}
        onClose={() => setDetail(null)}
        title="能力资产详情"
        extra={
          <Button
            type="primary"
            onClick={() => {
              setDetail(null);
              go("build");
            }}
          >
            在产品建设中调用
          </Button>
        }
      >
        {detail &&
          (() => {
            const info = detailInfo(detail);
            return (
              <>
                <div className="asset-detail-head">
                  <DeploymentUnitOutlined />
                  <div>
                    <Tag color="red">{detail.type}</Tag>
                    <h2>{detail.name}</h2>
                    <p>
                      {detail.id} · {detail.version}
                    </p>
                  </div>
                </div>
                <Card size="small" className="asset-detail-card">
                  <b>项目介绍</b>
                  <p>{info.purpose}</p>
                </Card>
                <Descriptions
                  bordered
                  size="small"
                  column={2}
                  items={[
                    { label: "来源产品", children: detail.source },
                    { label: "维护主体", children: detail.owner },
                    { label: "标准输入", children: detail.input },
                    { label: "标准输出", children: detail.output },
                    { label: "适用场景", children: info.scene, span: 2 },
                    { label: "调用方式", children: info.usage, span: 2 },
                    { label: "可配置参数", children: info.params, span: 2 },
                    { label: "配置边界", children: info.boundary, span: 2 },
                    { label: "关联情况", children: info.relations, span: 2 },
                    {
                      label: "运行状态",
                      children: <StatusTag status="正常运行" />,
                    },
                    { label: "当前版本", children: detail.version },
                  ]}
                />
                <Tabs
                  items={[
                    {
                      key: "use",
                      label: "使用步骤",
                      children: (
                        <Timeline
                          items={[
                            {
                              children: "在需求拆解或建设路径中识别可复用资产",
                            },
                            { children: `选择 ${detail.name} 并核对输入输出` },
                            { children: "配置允许调整的参数并完成样本验证" },
                            {
                              children: "随产品版本提交审批，正式建立调用关系",
                            },
                          ]}
                        />
                      ),
                    },
                    {
                      key: "relation",
                      label: "关联产品",
                      children: (
                        <List
                          size="small"
                          dataSource={
                            detail.calls
                              ? [
                                  detail.source,
                                  "浙江分行专精特新企业筛选产品",
                                  "产业事件驱动融资机会识别",
                                ].slice(0, Math.min(3, detail.calls))
                              : []
                          }
                          renderItem={(x: string) => (
                            <List.Item
                              actions={[
                                <Button
                                  key="view"
                                  type="link"
                                  onClick={() => go("finished")}
                                >
                                  查看产品
                                </Button>,
                              ]}
                            >
                              <ProductOutlined /> {x}
                            </List.Item>
                          )}
                          locale={{ emptyText: "刚刚回沉，尚未被其他产品调用" }}
                        />
                      ),
                    },
                    {
                      key: "version",
                      label: "版本记录",
                      children: (
                        <Timeline
                          items={[
                            { children: `${detail.version} 当前生效版本` },
                            { children: "完成标准输入输出校验" },
                            { children: "通过能力资产入库审核" },
                          ]}
                        />
                      ),
                    },
                  ]}
                />
              </>
            );
          })()}
      </Drawer>
      <Modal
        width={1120}
        open={builderOpen}
        title="能力资产手工建设工坊"
        onCancel={() => setBuilderOpen(false)}
        footer={
          <Space>
            <Button disabled={builderStep === 0} onClick={() => setBuilderStep(builderStep - 1)}>上一步</Button>
            {builderStep < 3 ? (
              <Button type="primary" disabled={builderStep === 2 && !builderTested} onClick={() => setBuilderStep(builderStep + 1)}>{builderStep === 2 ? "测试通过，进入治理" : "保存并继续"}</Button>
            ) : (
              <Button type="primary" onClick={() => {
                const draft = { id: builderKind === "module" ? "M-DRAFT-012" : "T-DRAFT-007", kind: builderKind, name: builderName, version: "V0.1", type: builderKind === "module" ? "能力模块" : "产品母版", source: "手工建设工坊", input: builderKind === "module" ? "event_id＋enterprise_id＋事件属性" : "产业事件＋企业主体＋参数", output: builderKind === "module" ? "影响评分＋方向＋证据" : "名单＋简报＋行动任务", calls: 0, owner: "总行数据管理部", status: "待入库审核" };
                localStorage.setItem("icbc-manual-asset-draft", JSON.stringify(draft));
                setManualAssets(old => [draft, ...old.filter(x=>x.id!==draft.id)]);
                setType(builderKind === "module" ? "module" : "template");
                message.success("能力资产包已生成并提交待入库审核");
                setBuilderOpen(false);
              }}>提交测试验证</Button>
            )}
          </Space>
        }
      >
        <Steps
          current={builderStep}
          items={["定义资产", "配置执行体", "样本测试", "版本与治理"].map(title => ({ title }))}
          style={{ marginBottom: 22 }}
        />
        {builderStep === 0 && <Row gutter={16}>
          <Col span={7}>
            <Card size="small" title="选择资产形态">
              <Radio.Group value={builderKind} onChange={e => { setBuilderKind(e.target.value); setBuilderName(e.target.value === "module" ? "产业事件影响评分" : "产业事件研判产品母版"); setBuilderTested(false); }} style={{ display: "grid", gap: 10 }}>
                <Radio.Button value="module">能力模块：可独立运行的加工单元</Radio.Button>
                <Radio.Button value="template">产品母版：完整产品结构蓝图</Radio.Button>
              </Radio.Group>
              <div className="asset-type-note" style={{ marginTop: 14 }}>
                <b>不是一份 Markdown 说明书</b>
                <span>说明文档只解释用途；真正可运行的资产还必须包含机器可读契约、执行逻辑、依赖、测试和治理元数据。</span>
              </div>
            </Card>
          </Col>
          <Col span={17}>
            <Card size="small" title={builderKind === "module" ? "模块基本定义" : "母版基本定义"}>
              <Form layout="vertical">
                <Row gutter={12}>
                  <Col span={12}><Form.Item label="名称"><Input value={builderName} onChange={e=>setBuilderName(e.target.value)} /></Form.Item></Col>
                  <Col span={6}><Form.Item label="资产编号"><Input defaultValue={builderKind === "module" ? "M012" : "T007"} /></Form.Item></Col>
                  <Col span={6}><Form.Item label="草稿版本"><Input defaultValue="V0.1" /></Form.Item></Col>
                </Row>
                <Form.Item label="业务目的"><Input.TextArea rows={2} defaultValue="根据事件类型、金额、主体关系和时效性，形成可解释的企业影响评分，供机会识别与风险核验产品调用。" /></Form.Item>
                <Row gutter={12}><Col span={12}><Form.Item label="标准输入"><Input defaultValue="event_id、enterprise_id、event_type、amount、event_date" /></Form.Item></Col><Col span={12}><Form.Item label="标准输出"><Input defaultValue="impact_score、direction、confidence、evidence_ids" /></Form.Item></Col></Row>
              </Form>
            </Card>
          </Col>
        </Row>}
        {builderStep === 1 && <Tabs items={builderKind === "module" ? [
          { key: "contract", label: "输入输出契约（JSON Schema）", children: <Input.TextArea rows={15} defaultValue={'{\n  "input": {\n    "event_id": "string|required",\n    "enterprise_id": "string|required",\n    "event_type": "enum:政策|项目|订单|处罚|司法",\n    "amount": "number|null",\n    "event_date": "date|required"\n  },\n  "output": {\n    "impact_score": "number:0-100",\n    "direction": "enum:机会|风险|中性",\n    "confidence": "number:0-1",\n    "evidence_ids": "string[]"\n  }\n}'} /> },
          { key: "logic", label: "执行逻辑（SQL/Python）", children: <Input.TextArea rows={15} defaultValue={'-- 演示性确定规则；生产环境可替换为审批后的模型服务\nSELECT enterprise_id,\n  LEAST(100, event_weight * freshness_weight * relation_weight) AS impact_score,\n  CASE WHEN event_type IN (\'处罚\',\'司法\') THEN \'风险\' ELSE \'机会\' END AS direction\nFROM standardized_events;'} /> },
          { key: "config", label: "参数与依赖（YAML）", children: <Input.TextArea rows={15} defaultValue={'runtime: batch-sql\ndependencies:\n  - M001_subject_linkage@2.3\n  - EXT_R021_standard_event@1.1\nparameters:\n  observation_days: 180\n  min_confidence: 0.70\nexception_policy:\n  missing_amount: keep_and_reduce_confidence\n  subject_unmatched: manual_review'} /> },
        ] : [
          { key: "flow", label: "流程清单", children: <Table pagination={false} dataSource={["事件输入","事件标准化","主体关联","影响评分","风险核验","人工复核","名单与简报输出"].map((name,i)=>({key:i,seq:i+1,name,asset:i===0||i===6?"输入/输出节点":`M00${Math.min(i,9)} 标准模块`,editable:i===5?"必须保留":"可替换同契约模块"}))} columns={[{title:"顺序",dataIndex:"seq"},{title:"流程节点",dataIndex:"name"},{title:"引用资产",dataIndex:"asset"},{title:"适配边界",dataIndex:"editable"}]} /> },
          { key: "manifest", label: "母版清单（YAML）", children: <Input.TextArea rows={15} defaultValue={'template_id: T007\nscenario: 产业事件研判\nflow: [event_input, normalize, subject_link, score, risk_check, human_review, deliver]\nrules: [region_scope, observation_window, risk_exclusion]\nmappings: [event_standard_v1, enterprise_master_v3]\ndeliverables: [enterprise_list.xlsx, industry_brief.docx, action_tasks.json]\nlocked: [subject_id_standard, evidence_trace, human_review]'} /> }
        ]} />}
        {builderStep === 2 && <Row gutter={16}>
          <Col span={15}><Card size="small" title="测试样本与预期结果" extra={<Button type="primary" icon={<ExperimentOutlined />} onClick={()=>{setBuilderTested(true);message.success("已执行3组测试：契约、计算结果与异常分支均符合预期");}}>实际运行测试</Button>}><Table pagination={false} dataSource={[
            {key:1,event:"某企业中标动力电池设备项目",amount:"2.8亿元",date:"2026-07-18",expected:"机会 / 82分",result:"通过"},
            {key:2,event:"某企业新增被执行事项",amount:"1260万元",date:"2026-07-26",expected:"风险 / 76分",result:"通过"},
            {key:3,event:"事件主体无法匹配",amount:"—",date:"2026-07-30",expected:"转人工复核",result:"通过"},
          ]} columns={[{title:"输入事件",dataIndex:"event"},{title:"金额",dataIndex:"amount"},{title:"日期",dataIndex:"date"},{title:"预期输出",dataIndex:"expected"},{title:"运行结果",dataIndex:"result",render:x=>builderTested?<Tag color="green">{x}</Tag>:<Tag>待运行</Tag>}]} /></Card></Col>
          <Col span={9}><Card size="small" title="验证门槛"><Descriptions column={1} bordered size="small" items={[{label:"契约校验",children:builderTested?<Tag color="green">字段、类型、必填项通过</Tag>:"待运行"},{label:"功能正确率",children:builderTested?"3/3 演示样本通过":"待运行"},{label:"异常处理",children:"未匹配主体进入人工复核"},{label:"可追溯性",children:"输出保留 evidence_ids"},{label:"当前结论",children:builderTested?<Tag color="blue">可进入联调，仍不可直接发布</Tag>:<Tag color="orange">测试未执行</Tag>}]} /></Card><div className="builder-test-note"><b>这里执行的是什么？</b><span>不是让大模型“读懂说明”，而是以样本数据调用确定性执行体，校验输入契约、计算输出、异常分支和证据链。</span></div></Col>
        </Row>}
        {builderStep === 3 && <Row gutter={16}>
          <Col span={14}><Card size="small" title="版本与责任信息"><Descriptions bordered column={2} items={[{label:"版本",children:"V0.1 待验证"},{label:"维护部门",children:"总行数据管理部"},{label:"业务责任部门",children:"公司金融业务部"},{label:"技术责任部门",children:"科技开发部"},{label:"适用范围",children:"验证环境 / 四川样本"},{label:"变更方式",children:"契约变更需新建大版本"},{label:"授权边界",children:"仅调用已授权字段",span:2},{label:"发布路径",children:"样本验证 → 技术联调 → 业务验收 → 审批入库",span:2}]} /></Card></Col>
          <Col span={10}><Card size="small" title="形成的资产包"><List dataSource={["manifest.yaml：资产元数据与依赖","input.schema.json / output.schema.json：机器可读契约","transform.sql 或 handler.py：执行逻辑","rules.yaml：可配置口径与边界","tests/：样本、预期结果与回归测试","README.md：业务说明和调用示例"]} renderItem={x=><List.Item><FileDoneOutlined /> {x}</List.Item>} /></Card></Col>
        </Row>}
      </Modal>
    </div>
  );
}

function FinishedProducts({
  published,
  productName,
  onUse,
}: {
  published: boolean;
  productName: string;
  onUse: (name: string) => void;
}) {
  const [offlineIds, setOfflineIds] = useState<string[]>([]);
  const [active, setActive] = useState("all");
  const [keyword, setKeyword] = useState("");
  const [detail, setDetail] = useState<any>(null);
  const [optimizedIds, setOptimizedIds] = useState<string[]>([]);
  const [businessCategory, setBusinessCategory] = useState("全部业务类别");
  const [region, setRegion] = useState("全部地域");
  const [runStatus, setRunStatus] = useState("全部运行状态");
  useEffect(
    () =>
      setOfflineIds(
        JSON.parse(localStorage.getItem("icbc-offline-products") || "[]"),
      ),
    [],
  );
  useEffect(
    () =>
      setOptimizedIds(
        JSON.parse(localStorage.getItem("icbc-optimized-products") || "[]"),
      ),
    [],
  );
  const localProducts = [
    {
      id: "P-ZJ-2026-012",
      name: "浙江先进制造企业成长画像",
      category: "产业研究",
      scope: "地方",
      users: "浙江分行",
      rating: 4.6,
      uses: 1186,
      status: "正常运行",
      version: "V1.7",
      owner: "浙江分行数据管理部",
    },
    {
      id: "P-SH-2026-018",
      name: "上海科创企业融资机会识别",
      category: "客户营销",
      scope: "地方",
      users: "上海分行",
      rating: 4.6,
      uses: 1860,
      status: "正常运行",
      version: "V2.1",
      owner: "上海分行数据管理部",
    },
    {
      id: "P-JS-2026-023",
      name: "江苏招投标客户拓展名单",
      category: "客户营销",
      scope: "地方",
      users: "江苏分行",
      rating: 4.5,
      uses: 1428,
      status: "正常运行",
      version: "V1.8",
      owner: "江苏分行数据管理部",
    },
    {
      id: "P-GD-2026-009",
      name: "广东供应链核心企业监测",
      category: "风险监测",
      scope: "地方",
      users: "广东分行",
      rating: 4.2,
      uses: 960,
      status: "待优化",
      version: "V1.4",
      owner: "广东分行数据管理部",
    },
    {
      id: "P-SC-2026-006",
      name: "四川重点项目融资线索",
      category: "客户营销",
      scope: "地方",
      users: "四川分行",
      rating: 4.4,
      uses: 815,
      status: "正常运行",
      version: "V1.5",
      owner: "四川分行数据管理部",
    },
    {
      id: "P-BJ-2026-015",
      name: "北京园区高新企业画像",
      category: "经营分析",
      scope: "地方",
      users: "北京分行",
      rating: 4.1,
      uses: 736,
      status: "待优化",
      version: "V1.2",
      owner: "北京分行数据管理部",
    },
    {
      id: "P-LD-2026-003",
      name: "长三角产业链关联风险图谱",
      category: "风险监测",
      scope: "限域",
      users: "沪苏浙皖分行",
      rating: 4.3,
      uses: 628,
      status: "正常运行",
      version: "V1.6",
      owner: "总行数据管理部",
    },
    {
      id: "P-LD-2026-007",
      name: "境内债券发行主体监测",
      category: "授信支持",
      scope: "限域",
      users: "金融市场相关机构",
      rating: 4.0,
      uses: 412,
      status: "待优化",
      version: "V1.1",
      owner: "总行金融市场部",
    },
    {
      id: "P-LD-2026-011",
      name: "核心客户跨境舆情预警",
      category: "风险监测",
      scope: "限域",
      users: "指定境外机构",
      rating: 4.5,
      uses: 358,
      status: "正常运行",
      version: "V2.0",
      owner: "总行风险管理部",
    },
  ];
  const resolveRegion = (x: any) =>
    x.scope === "全行"
      ? "全国"
      : String(x.users || "").includes("沪苏浙皖")
        ? "长三角"
        : String(x.users || "").replace(/省|市|分行/g, "") || "指定机构";
  const data = [
    ...(published
      ? [
          {
            id: "P-2026-089",
            name: productName,
            category: "客户营销",
            scope: "全行",
            users: "全行机构",
            rating: "—",
            uses: 0,
            status: "正常运行",
            version: "V3.0",
            owner: "总行数据管理部、公司金融部",
          },
        ]
      : []),
    ...initialProducts.map((x) => ({
      ...x,
      version: "V3.2",
      owner: "总行数据管理部",
    })),
    ...localProducts,
  ].map((x) => ({
    ...x,
    region: resolveRegion(x),
    status: optimizedIds.includes(x.id) ? "正常运行" : x.status,
  }));
  useEffect(() => {
    if (
      published &&
      data[0]?.id === "P-2026-089" &&
      localStorage.getItem("icbc-open-published-detail") === "true"
    ) {
      setDetail(data[0]);
      localStorage.removeItem("icbc-open-published-detail");
    }
  }, [published]);
  const tabs = [
    {
      key: "all",
      label: "全部成品",
      count: 86,
      desc: "已正式发布并纳入统一管理的外数产品",
      icon: <ProductOutlined />,
    },
    {
      key: "bank",
      label: "全行产品",
      count: 49,
      desc: "面向全行机构统一发布、统一维护",
      icon: <GlobalOutlined />,
    },
    {
      key: "local",
      label: "地方产品",
      count: 31 + (published ? 1 : 0),
      desc: "由分行建设并在特定地域使用",
      icon: <HomeOutlined />,
    },
    {
      key: "limited",
      label: "限域产品",
      count: 6,
      desc: "受数据授权或业务范围限制使用",
      icon: <SafetyCertificateOutlined />,
    },
    {
      key: "optimize",
      label: "待优化",
      count: Math.max(0, 7 - optimizedIds.length),
      desc: "评价或运行指标触发优化事项",
      icon: <ToolOutlined />,
    },
  ];
  const filtered = data.filter((p: any) => {
    const match =
      active === "all" ||
      (active === "bank" && p.scope === "全行") ||
      (active === "local" && p.scope === "地方") ||
      (active === "limited" && p.scope === "限域") ||
      (active === "optimize" && p.status === "待优化");
    const categoryMatch =
      businessCategory === "全部业务类别" || p.category === businessCategory;
    const regionMatch = region === "全部地域" || p.region === region;
    const displayedStatus = offlineIds.includes(p.id) ? "已下架" : p.status;
    const statusMatch =
      runStatus === "全部运行状态" || displayedStatus === runStatus;
    return (
      match &&
      categoryMatch &&
      regionMatch &&
      statusMatch &&
      (!keyword ||
        p.name.includes(keyword) ||
        p.id.toLowerCase().includes(keyword.toLowerCase()))
    );
  });
  const takeOffline = (product: any) =>
    Modal.confirm({
      title: "确认下架该产品？",
      content: (
        <div>
          <p>产品：{product.name}</p>
          <p className="muted">
            下架后，该产品将停止在应用门户提供服务，历史版本、运行记录和关联关系仍予保留。
          </p>
        </div>
      ),
      okText: "确认下架",
      cancelText: "取消",
      okButtonProps: { danger: true },
      onOk: () => {
        const next = [...new Set([...offlineIds, product.id])];
        setOfflineIds(next);
        localStorage.setItem("icbc-offline-products", JSON.stringify(next));
        message.success(`产品“${product.name}”已下架`);
      },
    });
  const finishOptimize = (product: any) =>
    Modal.confirm({
      title: "确认完成优化事项？",
      content: `产品“${product.name}”将恢复为正常运行状态，相关处理记录将保留。`,
      okText: "确认完成",
      onOk: () => {
        const next = [...new Set([...optimizedIds, product.id])];
        setOptimizedIds(next);
        localStorage.setItem("icbc-optimized-products", JSON.stringify(next));
        message.success("优化事项已完成，产品状态已更新");
      },
    });
  const manage = (key: string, product: any) => {
    if (key === "offline") return takeOffline(product);
    if (key === "detail") return setDetail(product);
    if (key === "optimize")
      return product.status === "待优化"
        ? finishOptimize(product)
        : message.success(`已为“${product.name}”创建优化事项`);
    const labels: any = {
      version: "新版本草稿已创建",
      scope: "适用范围调整申请已发起",
      source: "数据来源切换方案已生成",
      upgrade: "升级为全行产品评估已发起",
    };
    message.success(labels[key] || "操作已提交");
  };
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            成品管理区 <DemoTag />
          </h1>
          <p>
            统一管理正式产品的分类、版本、适用范围、运行状态、用户反馈和退出安排
          </p>
        </div>
        <Space>
          <Button onClick={() => message.success("成品清单已导出")}>
            导出清单
          </Button>
          <Button
            type="primary"
            onClick={() => message.success("新版本发布任务已创建")}
          >
            发布新版本
          </Button>
        </Space>
      </div>
      <div className="finished-tabs">
        {tabs.map((t) => (
          <button
            key={t.key}
            className={active === t.key ? "active" : ""}
            onClick={() => setActive(t.key)}
          >
            <span>{t.icon}</span>
            <div>
              <b>{t.label}</b>
              <small>{t.desc}</small>
            </div>
            <strong>{t.count}</strong>
          </button>
        ))}
      </div>
      <Row gutter={14} className="finished-insights">
        <Col span={6}>
          <Card>
            <Statistic title="本月新增发布" value={12} suffix="项" />
            <span>较上月增加 3 项</span>
          </Card>
        </Col>
        <Col span={6}>
          <Card>
            <Statistic title="覆盖使用机构" value={36} suffix="家" />
            <span>全行覆盖率 97.3%</span>
          </Card>
        </Col>
        <Col span={6}>
          <Card>
            <Statistic
              title="平均用户评分"
              value={4.56}
              precision={2}
              suffix="分"
            />
            <span>本月提升 0.08 分</span>
          </Card>
        </Col>
        <Col span={6}>
          <Card>
            <Statistic title="待处理维护事项" value={9} suffix="项" />
            <span className="red">其中逾期 2 项</span>
          </Card>
        </Col>
      </Row>
      <Card
        size="small"
        className="advanced-product-filter"
        title={
          <span>
            <SearchOutlined /> 高级筛选
          </span>
        }
        extra={
          <Button
            type="link"
            onClick={() => {
              setBusinessCategory("全部业务类别");
              setRegion("全部地域");
              setRunStatus("全部运行状态");
              setKeyword("");
            }}
          >
            重置筛选
          </Button>
        }
      >
        <div>
          <Select
            value={businessCategory}
            onChange={setBusinessCategory}
            options={[
              "全部业务类别",
              "客户营销",
              "风险监测",
              "授信支持",
              "产业研究",
              "经营分析",
              "监管与合规",
            ].map((value) => ({ value }))}
          />
          <Select
            value={region}
            onChange={setRegion}
            options={[
              "全部地域",
              "全国",
              "北京",
              "上海",
              "江苏",
              "浙江",
              "广东",
              "四川",
              "长三角",
              "指定机构",
            ].map((value) => ({ value }))}
          />
          <Select
            value={runStatus}
            onChange={setRunStatus}
            options={["全部运行状态", "正常运行", "待优化", "已下架"].map(
              (value) => ({ value }),
            )}
          />
          <Input
            value={keyword}
            onChange={(e) => setKeyword(e.target.value)}
            prefix={<SearchOutlined />}
            placeholder="产品名称或编号"
          />
        </div>
        <span>
          支持“业务分类 × 地域 × 运行状态”交叉筛选；当前匹配{" "}
          <b>{filtered.length}</b> 项
        </span>
      </Card>
      <Card
        className="finished-workspace"
        title={`${tabs.find((t) => t.key === active)?.label}产品清单`}
        extra={
          <span className="muted">当前展示 {filtered.length} 项演示产品</span>
        }
      >
        <div className="table-tools">
          <Space wrap>
            <Tag>{businessCategory}</Tag>
            <Tag>{region}</Tag>
            <Tag>{runStatus}</Tag>
          </Space>
          <Space>
            <Button onClick={() => message.info("已按使用频次从高到低排序")}>
              使用分析
            </Button>
            <Button onClick={() => setActive("optimize")}>查看优化事项</Button>
          </Space>
        </div>
        <Table
          dataSource={filtered}
          rowKey="id"
          pagination={{ pageSize: 8, showTotal: (n) => `共 ${n} 项` }}
          columns={[
            {
              title: "正式产品",
              dataIndex: "name",
              render: (x: string, r: any) => (
                <div>
                  <a onClick={() => setDetail(r)}>{x}</a>
                  {r.id === "P-2026-089" && <Tag color="red">刚刚发布</Tag>}
                  <small className="cell-sub">
                    {r.id} · {r.category}
                  </small>
                </div>
              ),
            },
            {
              title: "适用范围",
              dataIndex: "scope",
              render: (x: string) => (
                <Tag color={x === "地方" ? "orange" : "blue"}>{x}</Tag>
              ),
            },
            { title: "当前版本", dataIndex: "version" },
            { title: "使用机构", dataIndex: "users" },
            { title: "累计使用", dataIndex: "uses" },
            {
              title: "用户评分",
              dataIndex: "rating",
              render: (x: any) => (x === "—" ? "暂无" : x + " 分"),
            },
            {
              title: "运行状态",
              dataIndex: "status",
              render: (_: string, r: any) => (
                <StatusTag
                  status={offlineIds.includes(r.id) ? "已下架" : r.status}
                />
              ),
            },
            {
              title: "操作",
              render: (_: any, r: any) => (
                <Dropdown
                  menu={{
                    onClick: ({ key }) => manage(key, r),
                    items: [
                      { key: "detail", label: "查看详情" },
                      { key: "version", label: "发布新版本" },
                      {
                        key: "optimize",
                        label:
                          r.status === "待优化"
                            ? "完成优化事项"
                            : "创建优化事项",
                      },
                      { key: "scope", label: "调整适用范围" },
                      { key: "source", label: "更换数据来源" },
                      {
                        key: "upgrade",
                        label: "升级为全行产品",
                        disabled: r.scope === "全行",
                      },
                      { type: "divider" },
                      {
                        key: "offline",
                        label: offlineIds.includes(r.id)
                          ? "产品已下架"
                          : "下架产品",
                        danger: true,
                        disabled: offlineIds.includes(r.id),
                      },
                    ],
                  }}
                >
                  <Button size="small">管理</Button>
                </Dropdown>
              ),
            },
          ]}
        />
      </Card>
      <Drawer
        width={680}
        open={!!detail}
        onClose={() => setDetail(null)}
        title="成品详情与管理"
        extra={
          <Button
            type="primary"
            onClick={() => message.success("版本变更申请已提交")}
          >
            发起变更
          </Button>
        }
      >
        {detail && (
          <>
            <div className="finished-detail-head">
              <ProductOutlined />
              <div>
                <Space>
                  <Tag color={detail.scope === "地方" ? "orange" : "blue"}>
                    {detail.scope}产品
                  </Tag>
                  <StatusTag
                    status={
                      offlineIds.includes(detail.id) ? "已下架" : detail.status
                    }
                  />
                </Space>
                <h2>{detail.name}</h2>
                <p>
                  {detail.id} · 当前版本 {detail.version}
                </p>
              </div>
            </div>
            <Descriptions
              bordered
              size="small"
              column={2}
              items={[
                { label: "业务类别", children: detail.category },
                { label: "适用机构", children: detail.users },
                { label: "维护主体", children: detail.owner },
                { label: "累计使用", children: `${detail.uses} 次` },
                {
                  label: "用户评分",
                  children:
                    detail.rating === "—" ? "暂无评价" : `${detail.rating} 分`,
                },
                {
                  label: detail.id.startsWith("P-ZJ") ? "发布时间" : "最近更新",
                  children: detail.id.startsWith("P-ZJ")
                    ? "刚刚发布"
                    : "2026-07-30",
                },
              ]}
            />
            <Tabs
              className="finished-detail-tabs"
              items={[
                {
                  key: "run",
                  label: "运行情况",
                  children: detail.id.startsWith("P-ZJ") ? (
                    <div className="no-run-data">
                      <Empty
                        image={Empty.PRESENTED_IMAGE_SIMPLE}
                        description={
                          <>
                            <b>暂无运行数据</b>
                            <span>
                              产品刚刚发布，尚未产生有效调用；近30日使用、结果采纳率和系统可用率将在首个统计周期后生成。
                            </span>
                          </>
                        }
                      />
                    </div>
                  ) : (
                    <Row gutter={12}>
                      <Col span={8}>
                        <Card>
                          <Statistic
                            title="近30日使用"
                            value={Math.round(detail.uses / 12)}
                          />
                        </Card>
                      </Col>
                      <Col span={8}>
                        <Card>
                          <Statistic
                            title="结果采纳率"
                            value={76.8}
                            suffix="%"
                          />
                        </Card>
                      </Col>
                      <Col span={8}>
                        <Card>
                          <Statistic
                            title="系统可用率"
                            value={99.6}
                            suffix="%"
                          />
                        </Card>
                      </Col>
                    </Row>
                  ),
                },
                {
                  key: "relation",
                  label: "上下游关系",
                  children: (
                    <div className="finished-lineage">
                      <span>外数资源 4项</span>
                      <em>→</em>
                      <b>{detail.name}</b>
                      <em>→</em>
                      <span>业务流程 2个</span>
                    </div>
                  ),
                },
                {
                  key: "change",
                  label: "变更记录",
                  children: (
                    <Timeline
                      items={[
                        { children: `${detail.version} 发布 · 2026-07-18` },
                        { children: "更新司法风险数据映射 · 2026-06-25" },
                        { children: "完成季度运行评价 · 2026-06-10" },
                      ]}
                    />
                  ),
                },
              ]}
            />
            <Card title="快捷管理" className="quick-manage">
              <Space wrap>
                {detail.id === "P-2026-089" && (
                  <Button
                    type="primary"
                    icon={<ToolOutlined />}
                    onClick={() => onUse(detail.name)}
                  >
                    进入产品工作台运行
                  </Button>
                )}
                <Button onClick={() => manage("version", detail)}>
                  发布新版本
                </Button>
                <Button onClick={() => manage("scope", detail)}>
                  调整适用范围
                </Button>
                <Button onClick={() => manage("source", detail)}>
                  更换数据来源
                </Button>
                <Button onClick={() => manage("optimize", detail)}>
                  {detail.status === "待优化" ? "完成优化" : "创建优化事项"}
                </Button>
                <Button
                  danger
                  disabled={offlineIds.includes(detail.id)}
                  onClick={() => takeOffline(detail)}
                >
                  {offlineIds.includes(detail.id) ? "产品已下架" : "下架产品"}
                </Button>
              </Space>
            </Card>
          </>
        )}
      </Drawer>
    </div>
  );
}

function Operations({ newReview }: { newReview: boolean }) {
  const data = [
    ["专精特新企业营销名单", 5160, 4.8, 78, 99.6],
    ["企业司法风险监测", 4220, 4.6, 72, 99.8],
    ["舆情风险预警", 3690, 4.3, 61, 98.9],
    ["招投标客户发现", 2870, 4.5, 69, 99.4],
  ].map((x, i) => ({
    key: i,
    name: x[0],
    uses: x[1],
    rating: x[2],
    adopt: x[3],
    available: x[4],
  }));
  const latest = useMemo(() => {
    if (typeof window === "undefined") return null;
    try {
      return JSON.parse(localStorage.getItem("icbc-review-record") || "null");
    } catch {
      return null;
    }
  }, [newReview]);
  const baseFeedbacks = [
    {
      id: "FB-2026-028",
      product: "舆情风险预警",
      type: "结果解释不足",
      text: "舆情预警结果解释不足",
      priority: "高",
    },
    {
      id: "FB-2026-027",
      product: "企业司法风险监测",
      type: "新增需求",
      text: "司法风险事件原文链接缺失",
      priority: "中",
    },
    {
      id: "FB-2026-026",
      product: "招投标客户发现",
      type: "更新延迟",
      text: "招投标数据更新延迟",
      priority: "中",
    },
  ];
  const feedbacks = latest
    ? [
        {
          id: latest.id,
          product: latest.product,
          type: latest.type,
          text: latest.text,
          priority: "中",
        },
        ...baseFeedbacks,
      ]
    : baseFeedbacks;
  const [items, setItems] = useState<any[]>(() => {
    if (typeof window === "undefined") return [];
    try {
      return JSON.parse(
        localStorage.getItem("icbc-optimization-items") || "[]",
      );
    } catch {
      return [];
    }
  });
  const createItem = (f: any) => {
    if (items.some((x) => x.feedbackId === f.id)) {
      message.info("该反馈已形成优化事项");
      return;
    }
    const item = {
      id: `OPT-2026-${String(31 + items.length).padStart(3, "0")}`,
      feedbackId: f.id,
      product: f.product,
      title: f.text,
      owner: "待分派",
      due: "2026-08-12",
      status: "待分派",
    };
    const next = [item, ...items];
    setItems(next);
    localStorage.setItem("icbc-optimization-items", JSON.stringify(next));
    message.success(`已创建优化事项 ${item.id}`);
  };
  const finishItem = (id: string) => {
    const next = items.map((x) =>
      x.id === id ? { ...x, status: "已完成" } : x,
    );
    setItems(next);
    localStorage.setItem("icbc-optimization-items", JSON.stringify(next));
    message.success("优化事项已完成并保留处置记录");
  };
  const latestComment = latest
    ? {
        user: "浙江分行·郭经理",
        product: latest.product,
        score: latest.score || 5,
        text: latest.text,
        time: "刚刚",
      }
    : null;
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            运营评价 <DemoTag />
          </h1>
          <p>客观运行指标与主观使用评价共同驱动产品优化</p>
        </div>
        <Tag
          color={items.some((x) => x.status !== "已完成") ? "orange" : "green"}
        >
          优化事项 {items.length} 项
        </Tag>
      </div>
      <Row gutter={14}>
        <Col span={15}>
          <Card title="核心运营趋势">
            <ReactECharts
              style={{ height: 260 }}
              option={{
                tooltip: { trigger: "axis" },
                legend: { data: ["有效使用频次", "结果采纳率"] },
                xAxis: {
                  type: "category",
                  data: ["2月", "3月", "4月", "5月", "6月", "7月"],
                },
                yAxis: [{ type: "value" }, { type: "value", max: 100 }],
                series: [
                  {
                    name: "有效使用频次",
                    type: "bar",
                    data: [8600, 9100, 10200, 11800, 12600, 14400],
                    itemStyle: { color: "#c7000b" },
                  },
                  {
                    name: "结果采纳率",
                    type: "line",
                    yAxisIndex: 1,
                    data: [61, 64, 66, 68, 71, 74],
                    lineStyle: { color: "#d99000" },
                  },
                ],
              }}
            />
          </Card>
        </Col>
        <Col span={9}>
          <Card title="主观评价分布">
            <ReactECharts
              style={{ height: 260 }}
              option={{
                radar: {
                  indicator: [
                    "业务匹配",
                    "数据准确",
                    "数据时效",
                    "结果理解",
                    "操作便利",
                  ].map((name) => ({ name, max: 5 })),
                },
                series: [
                  {
                    type: "radar",
                    data: [{ value: [4.7, 4.5, 4.2, 4.6, 4.4] }],
                    areaStyle: { color: "rgba(199,0,11,.2)" },
                    lineStyle: { color: "#c7000b" },
                  },
                ],
              }}
            />
          </Card>
        </Col>
      </Row>
      <Card title="产品运营明细">
        <Table
          pagination={false}
          dataSource={data}
          columns={[
            { title: "产品名称", dataIndex: "name" },
            { title: "有效使用频次", dataIndex: "uses" },
            { title: "使用机构", render: () => "32 家" },
            {
              title: "结果采纳率",
              dataIndex: "adopt",
              render: (x: number) => x + "%",
            },
            {
              title: "系统可用率",
              dataIndex: "available",
              render: (x: number) => x + "%",
            },
            {
              title: "综合评分",
              dataIndex: "rating",
              render: (x: number) => (
                <>
                  <Rate disabled allowHalf value={x} />
                  <span>{x}</span>
                </>
              ),
            },
            { title: "维护成本", render: () => "中" },
            {
              title: "操作",
              render: () => <Button type="link">查看详情</Button>,
            },
          ]}
        />
      </Card>
      <Row gutter={14}>
        <Col span={14}>
          <Card title="最新用户评论">
            <List
              dataSource={latestComment ? [latestComment, ...reviews] : reviews}
              renderItem={(r) => (
                <List.Item>
                  <List.Item.Meta
                    avatar={<Avatar>{r.user.slice(-3, -2)}</Avatar>}
                    title={
                      <Space>
                        <b>{r.user}</b>
                        <Rate disabled value={r.score} />
                        <Tag>{r.product}</Tag>
                      </Space>
                    }
                    description={
                      <>
                        <p>{r.text}</p>
                        <small>{r.time}</small>
                      </>
                    }
                  />
                </List.Item>
              )}
            />
          </Card>
        </Col>
        <Col span={10}>
          <Card title="待处理反馈">
            <List
              dataSource={feedbacks}
              renderItem={(f: any, i) => (
                <List.Item
                  actions={[
                    <Button
                      key="a"
                      size="small"
                      disabled={items.some((x) => x.feedbackId === f.id)}
                      onClick={() => createItem(f)}
                    >
                      {items.some((x) => x.feedbackId === f.id)
                        ? "已创建优化项"
                        : "创建优化项"}
                    </Button>,
                  ]}
                >
                  <List.Item.Meta
                    title={
                      <Space>
                        <Badge status={i === 0 ? "error" : "warning"} />
                        <b>{f.text}</b>
                      </Space>
                    }
                    description={`${f.id} · ${f.product} · ${f.type}`}
                  />
                </List.Item>
              )}
            />
          </Card>
        </Col>
      </Row>
      {items.length > 0 && (
        <Card
          title="优化事项跟踪"
          extra={
            <Tag>
              {items.filter((x) => x.status !== "已完成").length} 项待处理
            </Tag>
          }
        >
          <Table
            pagination={false}
            rowKey="id"
            dataSource={items}
            columns={[
              { title: "事项编号", dataIndex: "id" },
              { title: "来源反馈", dataIndex: "feedbackId" },
              { title: "产品", dataIndex: "product" },
              { title: "优化内容", dataIndex: "title" },
              { title: "负责人", dataIndex: "owner" },
              { title: "完成期限", dataIndex: "due" },
              {
                title: "状态",
                dataIndex: "status",
                render: (x: string) => <StatusTag status={x} />,
              },
              {
                title: "操作",
                render: (_: any, r: any) => (
                  <Button
                    size="small"
                    type="primary"
                    disabled={r.status === "已完成"}
                    onClick={() => finishItem(r.id)}
                  >
                    {r.status === "已完成" ? "已完成" : "完成处置"}
                  </Button>
                ),
              },
            ]}
          />
        </Card>
      )}
    </div>
  );
}

function ValidationCenter({ go }: { go: (v: View) => void }) {
  const [running, setRunning] = useState(false);
  const [done, setDone] = useState(false);
  const start = () => {
    setRunning(true);
    setTimeout(() => {
      setRunning(false);
      setDone(true);
      message.success("6项测试用例执行完成，全部通过");
    }, 850);
  };
  const cases = [
    ["TC-F-001", "功能验证", "外数资源可正常获取"],
    ["TC-F-002", "功能验证", "主体关联准确率不低于98%"],
    ["TC-A-001", "适配验证", "地方名单完成主体ID映射"],
    ["TC-A-002", "适配验证", "浙江地区参数正确生效"],
    ["TC-R-001", "运行验证", "名单进入客户经理任务池"],
    ["TC-R-002", "运行验证", "营销反馈可回传运营评价"],
  ].map((x, i) => ({
    key: i,
    id: x[0],
    type: x[1],
    expect: x[2],
    actual: done
      ? [
          "资源返回完整",
          "样本准确率99.2%",
          "500条样本全部映射",
          "仅返回浙江省企业",
          "推送链路验证成功",
          "评价字段完整回传",
        ][i]
      : "等待执行",
    status: done ? "测试通过" : "待执行",
  }));
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            验证发布 <DemoTag />
          </h1>
          <p>集中完成产品功能、地方适配、业务运行验证及发布审批</p>
        </div>
        <Space>
          <Button onClick={() => message.success("测试方案已保存")}>
            保存方案
          </Button>
          <Button type="primary" loading={running} onClick={start}>
            开始全量测试
          </Button>
        </Space>
      </div>
      <Row gutter={14} className="metrics-grid">
        <Col span={6}>
          <MetricCard title="待验证产品" value={4} trend="其中逾期1项" alert />
        </Col>
        <Col span={6}>
          <MetricCard title="本月测试用例" value={36} trend="通过率94.4%" />
        </Col>
        <Col span={6}>
          <MetricCard title="待发布审批" value={2} trend="地方发布1项" />
        </Col>
        <Col span={6}>
          <MetricCard
            title="平均验证周期"
            value={3.2}
            suffix="天"
            trend="较上月缩短0.6天"
          />
        </Col>
      </Row>
      <Card
        title="当前验证任务"
        extra={<StatusTag status={done ? "测试通过" : "待验证"} />}
      >
        <Descriptions
          bordered
          size="small"
          column={4}
          items={[
            { label: "产品", children: "浙江分行专精特新企业筛选产品" },
            { label: "验证批次", children: "VAL-2026-0731-08" },
            { label: "责任人", children: "周毅" },
            { label: "计划完成", children: "2026-08-02" },
            { label: "建设路径", children: "母版适配＋能力复用" },
            { label: "适用范围", children: "浙江地区" },
            { label: "测试环境", children: "集成验证环境" },
            { label: "测试数据", children: "脱敏演示样本" },
          ]}
        />
      </Card>
      <Card title="测试用例与执行结果">
        <Table
          pagination={false}
          size="small"
          dataSource={cases}
          columns={[
            { title: "用例编号", dataIndex: "id" },
            {
              title: "验证类型",
              dataIndex: "type",
              render: (x: string) => <Tag>{x}</Tag>,
            },
            { title: "预期结果", dataIndex: "expect" },
            { title: "实际结果", dataIndex: "actual" },
            {
              title: "状态",
              dataIndex: "status",
              render: (x: string) => <StatusTag status={x} />,
            },
            {
              title: "操作",
              render: () => (
                <Button
                  size="small"
                  onClick={() => message.info("已打开测试证据与执行日志")}
                >
                  查看证据
                </Button>
              ),
            },
          ]}
        />
      </Card>
      {done && (
        <Card className="publish-note">
          <CheckCircleFilled />
          <div>
            <b>验证结论：通过</b>
            <p>功能、适配和运行验证均符合发布要求，可提交地方产品发布审批。</p>
          </div>
          <Button type="primary" onClick={() => go("build")}>
            进入发布准备
          </Button>
        </Card>
      )}
    </div>
  );
}

function ProcessCenter({ go }: { go: (v: View) => void }) {
  const [handled, setHandled] = useState<string[]>([]);
  const [active, setActive] = useState("todo");
  const [detail, setDetail] = useState<any>(null);
  const rows = [
    {
      id: "WF-260731-018",
      item: "浙江分行专精特新产品受理",
      node: "产品受理",
      owner: "郭泽宇",
      org: "浙江分行",
      limit: "今天17:00",
      status: "待处理",
    },
    {
      id: "WF-260731-015",
      item: "江苏招投标工具发布审批",
      node: "数据管理审核",
      owner: "王琳",
      org: "江苏分行",
      limit: "08-02",
      status: "待处理",
    },
    {
      id: "WF-260730-032",
      item: "广东供应链风险映射评审",
      node: "科技开发确认",
      owner: "陈嘉敏",
      org: "广东分行",
      limit: "08-01",
      status: "处理中",
    },
    {
      id: "WF-260729-021",
      item: "舆情预警V2.4优化事项",
      node: "运营优化",
      owner: "赵谦",
      org: "总行",
      limit: "08-04",
      status: "待分派",
    },
  ].filter((x) => !handled.includes(x.id));
  const started = [
    {
      id: "WF-260731-009",
      item: "浙江地方资质规则新增申请",
      node: "能力审核",
      owner: "任佳宁",
      org: "浙江分行",
      limit: "08-03",
      status: "处理中",
    },
    {
      id: "WF-260730-024",
      item: "产业事件产品V3.2变更申请",
      node: "测试验证",
      owner: "杜晨曦",
      org: "总行",
      limit: "08-05",
      status: "处理中",
    },
    {
      id: "WF-260729-017",
      item: "招投标数据源扩围申请",
      node: "数据管理审核",
      owner: "许文博",
      org: "总行",
      limit: "08-06",
      status: "待处理",
    },
  ];
  const doneRows = [
    {
      id: "WF-260728-012",
      item: "司法风险核验能力V2.3发布",
      node: "流程结束",
      owner: "梁思远",
      org: "总行",
      limit: "07-30",
      status: "已办结",
    },
    {
      id: "WF-260727-008",
      item: "专精特新营销母版参数调整",
      node: "流程结束",
      owner: "沈嘉禾",
      org: "总行",
      limit: "07-29",
      status: "已办结",
    },
  ];
  const shown =
    active === "todo"
      ? rows
      : active === "started"
        ? started
        : active === "done"
          ? doneRows
          : [...rows, ...started, ...doneRows];
  const openTarget = (r: any) =>
    r.node.includes("测试")
      ? go("build-testing")
      : r.item.includes("发布") || r.node.includes("审核")
        ? go("build-publish")
        : r.item.includes("舆情")
          ? go("operations")
          : r.item.includes("浙江分行")
            ? go("candidates")
            : go("build");
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            流程中心 <DemoTag />
          </h1>
          <p>统一查看分行报送受理、总行建设评审、测试发布和运营优化流程</p>
        </div>
        <Space>
          <Button>流程统计</Button>
          <Button type="primary" onClick={() => go("candidates")}>
            处理分行报送
          </Button>
        </Space>
      </div>
      <Row gutter={14} className="metrics-grid">
        <Col span={6}>
          <MetricCard
            title="我的待办"
            value={rows.length}
            trend="高优先级2项"
            alert
            onClick={() => setActive("todo")}
          />
        </Col>
        <Col span={6}>
          <MetricCard
            title="我发起的"
            value={12}
            trend="本月新增5项"
            onClick={() => setActive("started")}
          />
        </Col>
        <Col span={6}>
          <MetricCard
            title="本月办结"
            value={46}
            trend="按时办结率96%"
            onClick={() => setActive("done")}
          />
        </Col>
        <Col span={6}>
          <MetricCard
            title="平均处理时长"
            value={1.8}
            suffix="天"
            trend="较上月下降12%"
            onClick={() => setActive("all")}
          />
        </Col>
      </Row>
      <Card>
        <Tabs
          activeKey={active}
          onChange={setActive}
          items={[
            { key: "todo", label: `我的待办（${rows.length}）` },
            { key: "started", label: "我发起的（12）" },
            { key: "done", label: "已办事项（46）" },
            { key: "all", label: "全部流程" },
          ]}
        />
        <Table
          pagination={false}
          rowKey="id"
          dataSource={shown}
          columns={[
            {
              title: "流程编号",
              dataIndex: "id",
              render: (x: string, r: any) => (
                <a onClick={() => setDetail(r)}>{x}</a>
              ),
            },
            {
              title: "事项名称",
              dataIndex: "item",
              render: (x: string, r: any) => (
                <a onClick={() => setDetail(r)}>{x}</a>
              ),
            },
            { title: "当前节点", dataIndex: "node" },
            { title: "发起机构", dataIndex: "org" },
            { title: "负责人", dataIndex: "owner" },
            { title: "处理期限", dataIndex: "limit" },
            {
              title: "状态",
              dataIndex: "status",
              render: (x: string) => <StatusTag status={x} />,
            },
            {
              title: "操作",
              render: (_: any, r: any) => (
                <Space>
                  <Button size="small" onClick={() => openTarget(r)}>
                    {active === "done" ? "查看" : "办理"}
                  </Button>
                  {active === "todo" && (
                    <Button
                      size="small"
                      onClick={() => {
                        setHandled((v) => [...v, r.id]);
                        message.success("事项已转交并记录操作日志");
                      }}
                    >
                      转交
                    </Button>
                  )}
                </Space>
              ),
            },
          ]}
        />
      </Card>
      <Drawer
        width={640}
        open={!!detail}
        onClose={() => setDetail(null)}
        title="流程实例详情"
      >
        {detail && (
          <>
            <Descriptions
              bordered
              size="small"
              column={2}
              items={[
                { label: "流程编号", children: detail.id },
                {
                  label: "当前状态",
                  children: <StatusTag status={detail.status} />,
                },
                { label: "事项名称", children: detail.item, span: 2 },
                { label: "当前节点", children: detail.node },
                { label: "发起机构", children: detail.org },
                { label: "负责人", children: detail.owner },
                { label: "处理期限", children: detail.limit },
              ]}
            />
            <Card title="处理轨迹" size="small" className="process-history">
              <Timeline
                items={[
                  { children: "流程发起并完成材料校验" },
                  { children: "进入当前处理节点" },
                  { color: "gray", children: "等待后续审批或办理" },
                ]}
              />
            </Card>
            <Space>
              <Button onClick={() => setDetail(null)}>关闭</Button>
              <Button
                type="primary"
                onClick={() => {
                  setDetail(null);
                  openTarget(detail);
                }}
              >
                {active === "done" ? "查看关联业务" : "进入办理界面"}
              </Button>
            </Space>
          </>
        )}
      </Drawer>
    </div>
  );
}

function SystemConfig() {
  const [saved, setSaved] = useState(false);
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            系统配置 <DemoTag />
          </h1>
          <p>维护组织权限、业务字典、流程模板、提醒策略和基础参数</p>
        </div>
        <Button
          type="primary"
          onClick={() => {
            setSaved(true);
            message.success("系统配置已保存并生成变更记录");
          }}
        >
          保存全部配置
        </Button>
      </div>
      <Row gutter={14}>
        <Col span={6}>
          <Card className="config-nav">
            <Menu
              selectedKeys={["workflow"]}
              items={[
                { key: "org", icon: <TeamOutlined />, label: "组织与用户" },
                {
                  key: "auth",
                  icon: <SafetyCertificateOutlined />,
                  label: "角色与权限",
                },
                { key: "dict", icon: <TableOutlined />, label: "业务字典" },
                { key: "workflow", icon: <AuditOutlined />, label: "流程模板" },
                { key: "notice", icon: <BellOutlined />, label: "提醒策略" },
                { key: "param", icon: <SettingOutlined />, label: "系统参数" },
              ]}
            />
          </Card>
        </Col>
        <Col span={18}>
          <Card
            title="产品建设审批流程模板"
            extra={
              <Tag color={saved ? "green" : "blue"}>
                {saved ? "配置已保存" : "当前生效 V2.3"}
              </Tag>
            }
          >
            <Form layout="vertical">
              <Row gutter={16}>
                <Col span={12}>
                  <Form.Item label="模板名称">
                    <Input defaultValue="标准外数产品建设审批流程" />
                  </Form.Item>
                </Col>
                <Col span={12}>
                  <Form.Item label="适用产品类型">
                    <Select
                      mode="multiple"
                      defaultValue={["全行产品", "地方产品", "限域产品"]}
                      options={["全行产品", "地方产品", "限域产品"].map(
                        (value) => ({ value }),
                      )}
                    />
                  </Form.Item>
                </Col>
              </Row>
              <h3>审批节点</h3>
              {[
                "业务负责人确认",
                "数据管理部门审核",
                "科技开发确认",
                "产品发布审批",
              ].map((x, i) => (
                <div className="config-step" key={x}>
                  <span>{i + 1}</span>
                  <Input value={x} readOnly />
                  <Select
                    defaultValue={
                      [
                        "报送机构业务负责人",
                        "总行数据管理部",
                        "科技开发部门",
                        "产品管理委员会",
                      ][i]
                    }
                    options={[
                      {
                        value: [
                          "报送机构业务负责人",
                          "总行数据管理部",
                          "科技开发部门",
                          "产品管理委员会",
                        ][i],
                      },
                    ]}
                  />
                  <InputNumber
                    defaultValue={i === 3 ? 2 : 1}
                    addonAfter="工作日"
                  />
                  <Checkbox defaultChecked>必经</Checkbox>
                </div>
              ))}
              <Row gutter={16} style={{ marginTop: 20 }}>
                <Col span={12}>
                  <Form.Item label="逾期提醒">
                    <Select
                      defaultValue="到期前1天、到期当日"
                      options={[{ value: "到期前1天、到期当日" }]}
                    />
                  </Form.Item>
                </Col>
                <Col span={12}>
                  <Form.Item label="异常通知">
                    <Select
                      mode="multiple"
                      defaultValue={["当前负责人", "流程管理员"]}
                      options={["当前负责人", "流程管理员", "发起人"].map(
                        (value) => ({ value }),
                      )}
                    />
                  </Form.Item>
                </Col>
              </Row>
            </Form>
          </Card>
          <Card title="最近配置变更">
            <Timeline
              items={[
                { children: "流程模板V2.3生效 · 周毅 · 2026-07-28" },
                { children: "新增“限域产品”发布类型 · 王琳 · 2026-07-25" },
                { children: "调整能力回沉审核责任人 · 郭泽宇 · 2026-07-22" },
              ]}
            />
          </Card>
        </Col>
      </Row>
    </div>
  );
}

function AppHome({ go }: { go: (v: View) => void }) {
  const liveDate = useLiveChineseDate();
  return (
    <div>
      <div className="app-welcome">
        <div>
          <span>{liveDate}</span>
          <h1>下午好，欢迎使用外部数据产品服务</h1>
          <p>从业务任务出发，快速找到可直接使用的数据产品和外数资源。</p>
        </div>
        <DemoTag />
      </div>
      <Row gutter={[14, 14]}>
        <Col span={16}>
          <Card
            title="推荐产品"
            extra={<a onClick={() => go("app-products")}>查看全部</a>}
          >
            <div className="product-cards">
              {initialProducts.slice(0, 3).map((p, i) => (
                <button
                  key={p.id}
                  onClick={() => {
                    localStorage.setItem("icbc-selected-product", p.name);
                    go("product-detail");
                  }}
                >
                  <span className={"product-icon c" + i}>
                    {
                      [
                        <SolutionOutlined key="1" />,
                        <TeamOutlined key="2" />,
                        <SafetyCertificateOutlined key="3" />,
                      ][i]
                    }
                  </span>
                  <Tag color={i === 1 ? "red" : "default"}>{p.category}</Tag>
                  <h3>{p.name}</h3>
                  <p>{p.desc}</p>
                  <footer>
                    <span>
                      <Rate disabled value={p.rating} />
                      {p.rating}
                    </span>
                    <b>立即使用 →</b>
                  </footer>
                </button>
              ))}
            </div>
          </Card>
        </Col>
        <Col span={8}>
          <Card title="场景快捷入口">
            <div className="scene-grid">
              {[
                ["客户拓展", <TeamOutlined />],
                ["授信调查", <FileDoneOutlined />],
                ["风险监测", <SafetyCertificateOutlined />],
                ["产业研究", <LineChartOutlined />],
                ["经营分析", <TableOutlined />],
                ["监管合规", <AuditOutlined />],
              ].map((x) => (
                <button key={x[0] as string} onClick={() => go("scenes")}>
                  {x[1]}
                  <span>{x[0]}</span>
                </button>
              ))}
            </div>
          </Card>
        </Col>
        <Col span={9}>
          <Card
            title="我的关注"
            extra={<a onClick={() => go("follows")}>管理关注</a>}
          >
            <List
              size="small"
              dataSource={[
                "长三角高端装备产业链",
                "浙江专精特新企业名单更新",
                "12家存量客户出现司法风险变化",
                "新能源行业景气指标",
              ]}
              renderItem={(x, i) => (
                <List.Item
                  onClick={() => go("follows")}
                  className="follow-home-item"
                >
                  <Badge color={i === 2 ? "red" : "#c7000b"} text={x} />
                  <span className="muted">{i + 2} 条更新</span>
                </List.Item>
              )}
            />
          </Card>
        </Col>
        <Col span={9}>
          <Card title="近期更新">
            <Timeline
              items={[
                { children: "专精特新企业营销名单升级至 V3.2" },
                { children: "全国招投标公告库新增项目金额字段" },
                { children: "司法风险核验能力更新风险口径" },
              ]}
            />
          </Card>
        </Col>
        <Col span={6}>
          <Card title="使用记录">
            <Statistic title="本月产品使用" value={28} suffix="次" />
            <Progress percent={72} strokeColor="#c7000b" />
            <p className="muted">最常使用：专精特新企业营销名单</p>
          </Card>
        </Col>
      </Row>
    </div>
  );
}

function AppProducts({
  published,
  go,
}: {
  published: boolean;
  go: (v: View) => void;
}) {
  const [scope, setScope] = useState("全部范围");
  const [task, setTask] = useState("全部任务");
  const [shape, setShape] = useState("全部形态");
  const [method, setMethod] = useState("全部建设方式");
  const [query, setQuery] = useState("");
  const publishedRecord =
    typeof window !== "undefined"
      ? JSON.parse(localStorage.getItem("icbc-published-product") || "null")
      : null;
  const data = [
    ...(published
      ? [
          {
            id: publishedRecord?.id || "P-2026-089",
            name: publishedRecord?.name || "全国产业链机会与风险识别",
            category: "客户营销",
            scope: "全行",
            users: "对公客户经理、风险管理人员",
            rating: 5,
            uses: 0,
            status: "新发布",
            desc: "以产业事件为触发，融合企业基本面、外部信用和行内经营信息，动态更新机会、风险与客户关系判断。",
          },
        ]
      : []),
    ...initialProducts,
  ];
  const productShape = (p: any) =>
    p.category.includes("风险")
      ? "预警"
      : p.category.includes("研究") || p.category.includes("经营")
        ? "画像"
        : "名单";
  const productMethod = (p: any) =>
    String(p.scope).includes("全行")
      ? "标准产品"
      : String(p.scope).includes("限域")
        ? "能力组合"
        : "母版适配";
  const shown = data.filter(
    (p: any) =>
      (scope === "全部范围" ||
        (scope === "全行产品" && String(p.scope).includes("全行")) ||
        (scope === "地方产品" &&
          !String(p.scope).includes("全行") &&
          !String(p.scope).includes("限域")) ||
        (scope === "限域产品" && String(p.scope).includes("限域"))) &&
      (task === "全部任务" || p.category === task) &&
      (shape === "全部形态" || productShape(p) === shape) &&
      (method === "全部建设方式" || productMethod(p) === method) &&
      `${p.name}${p.desc}${p.category}`
        .toLowerCase()
        .includes(query.trim().toLowerCase()),
  );
  const openProduct = (p: any) => {
    localStorage.setItem("icbc-selected-product", p.name);
    go("product-detail");
  };
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            外数产品 <DemoTag />
          </h1>
          <p>
            以业务任务为主线，并按产品层级、适用地域、交付形态和建设方式进行多维筛选。
          </p>
        </div>
      </div>
      <Card className="product-filters">
        <Row gutter={12}>
          <Col span={5}>
            <Select
              value={scope}
              onChange={setScope}
              style={{ width: "100%" }}
              options={["全部范围", "全行产品", "地方产品", "限域产品"].map(
                (value) => ({ value }),
              )}
            />
          </Col>
          <Col span={5}>
            <Select
              value={task}
              onChange={setTask}
              style={{ width: "100%" }}
              options={[
                "全部任务",
                "客户营销",
                "授信支持",
                "风险监测",
                "产业研究",
                "经营分析",
                "监管与合规",
              ].map((value) => ({ value }))}
            />
          </Col>
          <Col span={5}>
            <Select
              value={shape}
              onChange={setShape}
              style={{ width: "100%" }}
              options={["全部形态", "名单", "画像", "预警"].map((value) => ({
                value,
              }))}
            />
          </Col>
          <Col span={5}>
            <Select
              value={method}
              onChange={setMethod}
              style={{ width: "100%" }}
              options={["全部建设方式", "标准产品", "母版适配", "能力组合"].map(
                (value) => ({ value }),
              )}
            />
          </Col>
          <Col span={4}>
            <Input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              prefix={<SearchOutlined />}
              placeholder="产品名称"
              allowClear
            />
          </Col>
        </Row>
        <div className="filter-result">
          <span>
            已找到 <b>{shown.length}</b> 个产品
          </span>
          {(scope !== "全部范围" ||
            task !== "全部任务" ||
            shape !== "全部形态" ||
            method !== "全部建设方式" ||
            query) && (
            <Button
              type="link"
              size="small"
              onClick={() => {
                setScope("全部范围");
                setTask("全部任务");
                setShape("全部形态");
                setMethod("全部建设方式");
                setQuery("");
              }}
            >
              清空筛选
            </Button>
          )}
        </div>
      </Card>
      <div className="category-chips">
        {[
          "全部任务",
          "客户营销",
          "授信支持",
          "风险监测",
          "产业研究",
          "经营分析",
          "监管与合规",
        ].map((x) => (
          <Button
            key={x}
            type={task === x ? "primary" : "default"}
            onClick={() => setTask(x)}
          >
            {x}
          </Button>
        ))}
      </div>
      {shown.length ? (
        <div className="app-product-grid">
          {shown.map((p: any) => (
            <Card key={p.id} hoverable onClick={() => openProduct(p)}>
              <div className="product-top">
                <span className="product-icon c1">
                  <ProductOutlined />
                </span>
                <Space>
                  <Tag color="red">{p.category}</Tag>
                  <Tag>
                    {String(p.scope).includes("全行")
                      ? "全行产品"
                      : String(p.scope).includes("限域")
                        ? "限域产品"
                        : "地方产品"}
                  </Tag>
                  {p.status === "新发布" && <Tag color="green">刚刚发布</Tag>}
                </Space>
              </div>
              <h2>{p.name}</h2>
              <p>{p.desc}</p>
              <dl>
                <dt>适用地域</dt>
                <dd>{p.scope}</dd>
                <dt>产品形态</dt>
                <dd>{productShape(p)}＋行动提示</dd>
                <dt>建设方式</dt>
                <dd>{productMethod(p)}</dd>
                <dt>累计使用</dt>
                <dd>{p.uses} 次</dd>
              </dl>
              <footer>
                <span>
                  <Rate disabled value={Number(p.rating)} />
                  {p.rating}
                </span>
                <Button
                  type="primary"
                  onClick={(e) => {
                    e.stopPropagation();
                    openProduct(p);
                  }}
                >
                  查看并使用
                </Button>
              </footer>
            </Card>
          ))}
        </div>
      ) : (
        <Card>
          <Empty description="未找到符合条件的产品">
            <Button
              type="primary"
              onClick={() => {
                setScope("全部范围");
                setTask("全部任务");
                setShape("全部形态");
                setMethod("全部建设方式");
                setQuery("");
              }}
            >
              清空筛选条件
            </Button>
          </Empty>
        </Card>
      )}
    </div>
  );
}

function IndustryCommandPreview() {
  const [focus, setFocus] = useState("设备更新政策");
  const [region, setRegion] = useState("全国");
  const [industry, setIndustry] = useState("动力电池");
  const [chain, setChain] = useState("全产业链");
  const [window, setWindow] = useState("近90日");
  const [generated, setGenerated] = useState(true);
  const [companyOpen, setCompanyOpen] = useState<any>(null);
  const [taskState, setTaskState] = useState<Record<string, string>>({
    "T-260801-01": "待触达",
    "T-260801-02": "授信核验中",
    "T-260801-03": "风险复核中",
    "T-260801-04": "已完成",
  });
  const [evidenceOpen, setEvidenceOpen] = useState(false);
  const eventMap: Record<string, any> = {
    设备更新政策: {
      summary:
        "设备更新政策加速先进制造企业技改投资，预计带动设备融资、项目贷款与供应链结算需求。",
      opportunity: "12.6亿元",
      risk: "原材料价格上涨可能挤压6家下游企业毛利",
      tasks: "已生成28项协同任务",
      nodes: [82, 46, 67, 38],
    },
    重大项目中标: {
      summary:
        "链主企业新获重大项目，订单向核心零部件、工业软件与物流服务环节传导。",
      opportunity: "8.4亿元",
      risk: "3家二级供应商存在履约能力待核验事项",
      tasks: "已生成19项协同任务",
      nodes: [76, 52, 71, 43],
    },
    关键材料涨价: {
      summary:
        "关键材料价格连续四周上涨，成本压力沿上游材料—核心部件—整机制造链条传导。",
      opportunity: "2.1亿元",
      risk: "14家存量客户触发利润与现金流压力预警",
      tasks: "已生成14项风险任务",
      nodes: [41, 78, 49, 72],
    },
  };
  const current = eventMap[focus];
  const graph = {
    tooltip: {},
    series: [
      {
        type: "graph",
        layout: "force",
        roam: true,
        label: { show: true, fontSize: 10 },
        force: { repulsion: 260, edgeLength: 105 },
        data: [
          {
            name: "政策/市场事件",
            symbolSize: 58,
            itemStyle: { color: "#c7000b" },
          },
          {
            name: "链主整机企业",
            symbolSize: 52,
            itemStyle: { color: "#8f1018" },
          },
          {
            name: "核心零部件",
            symbolSize: 45,
            itemStyle: { color: "#d77c00" },
          },
          { name: "工业软件", symbolSize: 40, itemStyle: { color: "#3b82f6" } },
          {
            name: "材料供应商",
            symbolSize: 42,
            itemStyle: { color: "#6b7280" },
          },
          {
            name: "物流与服务",
            symbolSize: 38,
            itemStyle: { color: "#10b981" },
          },
        ],
        links: [
          { source: "政策/市场事件", target: "链主整机企业" },
          { source: "链主整机企业", target: "核心零部件" },
          { source: "链主整机企业", target: "工业软件" },
          { source: "核心零部件", target: "材料供应商" },
          { source: "链主整机企业", target: "物流与服务" },
        ],
        lineStyle: { color: "#aab2bd", width: 2, curveness: 0.08 },
      },
    ],
  };
  const quadrant = {
    grid: { left: 42, right: 18, top: 20, bottom: 38 },
    xAxis: { name: "经营机会", min: 0, max: 100 },
    yAxis: { name: "风险压力", min: 0, max: 100 },
    tooltip: { trigger: "item", formatter: "{b}<br/>机会 {c0}<br/>风险 {c1}" },
    series: [
      {
        type: "scatter",
        symbolSize: (v: number[]) => 18 + v[2] / 5,
        data: [
          {
            name: "川南精密制造",
            value: [current.nodes[0], current.nodes[1], 92],
          },
          {
            name: "西部新能源材料",
            value: [current.nodes[2], current.nodes[3], 78],
          },
          { name: "蜀源工业软件", value: [71, 24, 66] },
          { name: "成渝物流科技", value: [54, 58, 60] },
        ],
        itemStyle: {
          color: (p: any) => (p.value[1] > 60 ? "#d77c00" : "#c7000b"),
        },
        label: { show: true, position: "top", formatter: "{b}", fontSize: 9 },
      },
    ],
  };
  const tasks = [
    {
      id: "T-260801-01",
      company: "川南精密制造",
      role: "客户经理",
      action: "核实设备更新计划及融资需求",
      due: "今日",
    },
    {
      id: "T-260801-02",
      company: "绵阳功率半导体",
      role: "授信人员",
      action: "核验新增产线、订单与现金流",
      due: "8月8日",
    },
    {
      id: "T-260801-03",
      company: "泸州绿色化工",
      role: "风险人员",
      action: "评估环保整改对授信敞口影响",
      due: "今日",
    },
    {
      id: "T-260801-04",
      company: "德阳航空零部件",
      role: "客户经理",
      action: "完成订单融资方案初步沟通",
      due: "已反馈",
    },
  ];
  const updateTask = (id: string) => {
    const nextState =
      taskState[id] === "待触达"
        ? "已触达"
        : taskState[id]?.includes("中")
          ? "已完成"
          : "已完成";
    setTaskState({ ...taskState, [id]: nextState });
    localStorage.setItem(
      "icbc-command-feedback",
      JSON.stringify({ ...taskState, [id]: nextState }),
    );
    message.success("任务状态已更新，结果已回流产品评价");
  };
  const companyRows = Array.from({ length: 80 }, (_, i) => {
    const seed = tasks[i % tasks.length];
    const cities = ["成都", "绵阳", "德阳", "宜宾", "苏州", "上海", "深圳", "合肥", "宁波", "武汉"];
    const segments = ["电芯制造", "上游材料", "核心零部件", "电池设备", "储能集成", "回收利用"];
    const companyTypes = ["精密制造", "新能源材料", "智能装备", "工业软件", "绿色科技", "供应链服务"];
    const opportunity = Math.max(58, 95 - ((i * 7) % 36));
    const risk = 18 + ((i * 11) % 61);
    return {
      ...seed,
      id: `ENT-${String(i + 1).padStart(3, "0")}`,
      company: i < tasks.length ? seed.company : `${cities[i % cities.length]}${companyTypes[i % companyTypes.length]}示范企业${String(i + 1).padStart(2, "0")}号`,
      region: cities[i % cities.length],
      chain: segments[i % segments.length],
      opportunity,
      risk,
      level: risk > 65 ? "审慎核验" : opportunity >= 85 ? "重点推荐" : opportunity >= 75 ? "优先营销" : "持续关注",
      basis: ["扩产备案", "新增订单", "招标活跃", "专精特新资质", "结算增长", "产业政策支持"].slice(0, 3 + (i % 3)).join("、"),
      action: risk > 65 ? "核验经营与司法风险后确定业务策略" : seed.action,
    };
  });
  const exportChart = () => {
    const canvas = document.querySelector(
      ".industry-command canvas",
    ) as HTMLCanvasElement | null;
    if (!canvas) {
      message.warning("图表尚未完成渲染");
      return;
    }
    const a = document.createElement("a");
    a.href = canvas.toDataURL("image/png");
    a.download = `${industry}_${region}_产业链分析图_演示.png`;
    a.click();
    message.success("图表图片已导出");
  };
  return (
    <div className="industry-command command-full">
      <div className="page-heading">
        <div>
          <h1>
            全国产业链机会与风险识别 <DemoTag />
          </h1>
          <p>全国标准产品 · 参数筛选—综合研判—事实链与传导—重点企业—业务行动</p>
        </div>
        <Space>
          <Button icon={<DownloadOutlined />} onClick={exportChart}>导出图表图片</Button>
          <Button onClick={() => { localStorage.setItem("icbc-national-plan", JSON.stringify({ region, industry, chain, focus, window })); message.success("已保存为可再次运行的产品方案"); }}>保存方案</Button>
          <Button
            type="primary"
            onClick={() => message.success("业务反馈已提交")}
          >
            提交反馈
          </Button>
        </Space>
      </div>
      <Card size="small" className="national-filter" title="运行参数" extra={<Tag color={generated ? "green" : "orange"}>{generated ? "研判已更新" : "参数已变化"}</Tag>}>
        <Space wrap>
          <Select value={region} onChange={(v) => { setRegion(v); setGenerated(false); }} style={{ width: 130 }} options={["全国","四川省","北京市","上海市","江苏省","浙江省","广东省"].map(value => ({ value }))} />
          <Select value={industry} onChange={(v) => { setIndustry(v); setGenerated(false); }} style={{ width: 150 }} options={["动力电池","新能源汽车","高端装备","集成电路","光伏","储能","机器人","生物医药"].map(value => ({ value }))} />
          <Select value={chain} onChange={(v) => { setChain(v); setGenerated(false); }} style={{ width: 140 }} options={["全产业链","上游材料","核心部件","整机/系统","回收与服务"].map(value => ({ value }))} />
          <Select value={focus} onChange={(v) => { setFocus(v); setGenerated(false); }} style={{ width: 150 }} options={Object.keys(eventMap).map(value => ({ value }))} />
          <Select value={window} onChange={(v) => { setWindow(v); setGenerated(false); }} style={{ width: 120 }} options={["近30日","近90日","近半年","近一年"].map(value => ({ value }))} />
          <Select defaultValue="全部企业" style={{ width: 120 }} options={["全部企业","存量客户","潜在客户","专精特新企业"].map(value => ({ value }))} />
          <Select defaultValue="机会≥70" style={{ width: 110 }} options={["机会≥70","机会≥80","机会≥90"].map(value => ({ value }))} />
          <Select defaultValue="风险≤60" style={{ width: 110 }} options={["风险≤60","风险≤40","排除重大风险"].map(value => ({ value }))} />
          <Select defaultValue="推荐20家" style={{ width: 110 }} options={["推荐20家","推荐50家","推荐100家"].map(value => ({ value }))} />
          <Button type="primary" icon={<ExperimentOutlined />} onClick={() => { setGenerated(true); localStorage.setItem("icbc-national-last-run", JSON.stringify({ region, industry, chain, focus, window })); message.success("图表、综合判断与名单已同步更新"); }}>生成综合研判</Button>
        </Space>
      </Card>
      <div className="command-toolbar">
        <div>
          <Tag color="red">总行标准产品 V3.1</Tag>
          <h2>{region} · {industry} · {chain}</h2>
          <p>
            从事件出发，同时回答影响哪条产业链、涉及哪些客户、机会和风险如何传导、业务应采取什么行动。
          </p>
        </div>
        <Select
          value={focus}
          onChange={setFocus}
          style={{ width: 190 }}
          options={Object.keys(eventMap).map((value) => ({ value }))}
        />
      </div>
      <div className="command-kpis">
        {[
          ["机会融资规模", current.opportunity],
          ["重点客户", "36家"],
          ["风险传导", current.risk],
          ["协同执行", current.tasks],
        ].map((x) => (
          <div key={x[0]}>
            <small>{x[0]}</small>
            <b>{x[1]}</b>
          </div>
        ))}
      </div>
      <Card
        size="small"
        className="event-evidence"
        title={
          <Space>
            <RobotOutlined />
            AI事件解析与证据链
          </Space>
        }
        extra={
          <Button type="link" onClick={() => setEvidenceOpen(!evidenceOpen)}>
            {evidenceOpen ? "收起证据" : "查看9项证据"}
          </Button>
        }
      >
        <div className="event-parse">
          <span>
            <small>事件类型</small>
            <b>{focus}</b>
          </span>
          <span>
            <small>影响方向</small>
            <b>设备投资与订单需求上升</b>
          </span>
          <span>
            <small>影响期限</small>
            <b>6—18个月</b>
          </span>
          <span>
            <small>解析置信度</small>
            <b>94%</b>
          </span>
        </div>
        {evidenceOpen && (
          <Timeline
            items={[
              {
                color: "red",
                children: "国家及省级设备更新政策原文 · 2026-07-30 · 权威等级A",
              },
              {
                color: "green",
                children:
                  "产业园重大项目备案与招标公告 · 2026-07-31 · 主体匹配98%",
              },
              {
                color: "blue",
                children: "链主企业产能、招聘及供应商订单变化 · 近30日持续核验",
              },
            ]}
          />
        )}
      </Card>
      <Row gutter={12}>
        <Col span={13}>
          <Card
            size="small"
            title="产业事件影响与链路传导图"
            extra={<Tag>事件切换实时联动</Tag>}
          >
            <ReactECharts style={{ height: 310 }} option={graph} />
          </Card>
        </Col>
        <Col span={11}>
          <Card size="small" title="链上客户机会—风险分布">
            <ReactECharts style={{ height: 310 }} option={quadrant} />
          </Card>
        </Col>
      </Row>
      <Row gutter={12}>
        <Col span={14}>
          <Card size="small" title="企业变化趋势与融资需求">
            <ReactECharts
              style={{ height: 260 }}
              option={{
                tooltip: { trigger: "axis" },
                legend: { data: ["机会指数", "风险指数", "融资需求"] },
                xAxis: {
                  type: "category",
                  data: ["3月", "4月", "5月", "6月", "7月", "8月"],
                },
                yAxis: [
                  { type: "value", max: 100 },
                  { type: "value", axisLabel: { formatter: "{value}亿" } },
                ],
                series: [
                  {
                    name: "机会指数",
                    type: "line",
                    smooth: true,
                    data: [52, 55, 61, 68, 74, 86],
                    itemStyle: { color: "#c7000b" },
                  },
                  {
                    name: "风险指数",
                    type: "line",
                    data: [29, 31, 30, 36, 42, 38],
                    itemStyle: { color: "#d98700" },
                  },
                  {
                    name: "融资需求",
                    type: "bar",
                    yAxisIndex: 1,
                    data: [1.8, 2.1, 2.6, 4.2, 7.5, 12.6],
                    itemStyle: { color: "#e8b6b9" },
                  },
                ],
              }}
            />
          </Card>
        </Col>
        <Col span={10}>
          <Card size="small" title="区域机会热力">
            <ReactECharts
              style={{ height: 260 }}
              option={{
                tooltip: {},
                xAxis: {
                  type: "category",
                  data: ["成都", "绵阳", "德阳", "宜宾", "眉山", "乐山"],
                },
                yAxis: {
                  type: "category",
                  data: ["整机", "零部件", "材料", "软件"],
                },
                visualMap: {
                  min: 20,
                  max: 95,
                  orient: "horizontal",
                  left: "center",
                  bottom: 0,
                  inRange: { color: ["#fff1f1", "#e36b72", "#a80712"] },
                },
                series: [
                  {
                    type: "heatmap",
                    data: [
                      [0, 0, 88],
                      [1, 0, 61],
                      [2, 0, 72],
                      [3, 0, 48],
                      [4, 0, 39],
                      [5, 0, 33],
                      [0, 1, 74],
                      [1, 1, 92],
                      [2, 1, 86],
                      [3, 1, 54],
                      [4, 1, 49],
                      [5, 1, 41],
                      [0, 2, 55],
                      [1, 2, 68],
                      [2, 2, 64],
                      [3, 2, 91],
                      [4, 2, 70],
                      [5, 2, 82],
                      [0, 3, 89],
                      [1, 3, 78],
                      [2, 3, 52],
                      [3, 3, 44],
                      [4, 3, 38],
                      [5, 3, 31],
                    ],
                  },
                ],
              }}
            />
          </Card>
        </Col>
      </Row>
      <Card
        size="small"
        className="ai-brief"
        title={
          <Space>
            <RobotOutlined />
            经营分析简报
          </Space>
        }
        extra={<Tag color="green">已核验外数证据 9 项</Tag>}
      >
        <p>{current.summary}</p>
        <div>
          {[
            "优先联系12家高机会低风险客户",
            "为4家链上企业生成设备更新融资方案",
            "对风险传导客户补充现金流与订单核验",
            "将结果分别路由客户经理、授信与风险岗位",
          ].map((x, i) => (
            <span key={x}>
              <b>0{i + 1}</b>
              {x}
            </span>
          ))}
        </div>
      </Card>
      <Card title="候选企业全景与分层明细" extra={<Space><Tag color="red">80家候选企业</Tag><Button icon={<DownloadOutlined />} onClick={() => exportExcel(companyRows, `${industry}_${region}_重点企业明细_演示.xls`)}>导出完整Excel企业明细</Button></Space>}>
        <Table size="small" scroll={{ x: 1500, y: 430 }} pagination={{ pageSize: 10, showSizeChanger: true, pageSizeOptions: [10,20,50], showTotal: total => `共 ${total} 家候选企业` }} rowKey="id" dataSource={companyRows} columns={[
          { title:"企业名称", dataIndex:"company", render:(x:string,r:any)=><a onClick={() => setCompanyOpen(r)}>{x}</a> },
          { title:"地区", dataIndex:"region" }, { title:"产业链环节", dataIndex:"chain" },
          { title:"机会评分", dataIndex:"opportunity", render:(x:number)=><Progress percent={x} size="small" strokeColor="#c7000b" /> },
          { title:"风险评分", dataIndex:"risk", render:(x:number)=><Tag color={x>60?"red":x>40?"orange":"green"}>{x}</Tag> },
          { title:"推荐层级", dataIndex:"level", render:(x:string)=><Tag color="red">{x}</Tag> },
          { title:"入选依据", dataIndex:"basis" }, { title:"建议业务方向", dataIndex:"action" },
          { title:"操作", render:(_:any,r:any)=><Button size="small" onClick={() => setCompanyOpen(r)}>企业详情</Button> }
        ]} />
      </Card>
      <Drawer width={720} open={!!companyOpen} onClose={() => setCompanyOpen(null)} title="企业机会与风险详情">
        {companyOpen && <><div className="finished-detail-head"><BankOutlined/><div><Tag color="red">{companyOpen.level}</Tag><h2>{companyOpen.company}</h2><p>{companyOpen.region} · {industry} · {companyOpen.chain}</p></div></div><Descriptions bordered column={2} items={[
          {label:"企业基本情况",children:"存续 · 制造业 · 演示企业"},{label:"产业链环节",children:companyOpen.chain},
          {label:"机会评分",children:`${companyOpen.opportunity}分`},{label:"风险评分",children:`${companyOpen.risk}分`},
          {label:"主要机会信号",children:companyOpen.basis,span:2},{label:"风险事件",children:companyOpen.risk>60?"存在需核验的司法或经营风险":"未见重大风险，持续监测",span:2},
          {label:"评分依据",children:"事件强度30%＋产业地位25%＋经营活跃20%＋客户关系15%＋风险调整10%",span:2},
          {label:"数据来源",children:"全国企业库、产业链图谱、工商股权、招投标、融资事件、司法风险及行内客户数据",span:2},
          {label:"建议动作",children:companyOpen.action,span:2},{label:"待核验事项",children:"订单真实性、资金用途、关联交易与最新司法事项",span:2}
        ]}/><div className="pane-actions"><Button onClick={() => message.success("已加入营销核验任务")}>发起营销动作</Button><Button onClick={() => message.success("已加入授信核验清单")}>发起授信核验</Button><Button type="primary" onClick={() => message.success("已加入风险持续监测")}>加入风险监测</Button></div></>}
      </Drawer>
      <Card
        title="跨岗位协同任务与结果回流"
        extra={
          <Space>
            <Tag>客户经理 12</Tag>
            <Tag>授信 9</Tag>
            <Tag>风险 7</Tag>
          </Space>
        }
      >
        <Table
          size="small"
          pagination={false}
          rowKey="id"
          dataSource={tasks}
          columns={[
            { title: "任务编号", dataIndex: "id" },
            {
              title: "企业",
              dataIndex: "company",
              render: (x: string) => <b>{x}</b>,
            },
            {
              title: "责任岗位",
              dataIndex: "role",
              render: (x: string) => (
                <Tag
                  color={
                    x === "风险人员"
                      ? "orange"
                      : x === "授信人员"
                        ? "blue"
                        : "red"
                  }
                >
                  {x}
                </Tag>
              ),
            },
            { title: "业务动作", dataIndex: "action" },
            { title: "期限", dataIndex: "due" },
            {
              title: "状态",
              render: (_: any, r: any) => (
                <StatusTag status={taskState[r.id]} />
              ),
            },
            {
              title: "操作",
              render: (_: any, r: any) => (
                <Button
                  size="small"
                  disabled={taskState[r.id] === "已完成"}
                  onClick={() => updateTask(r.id)}
                >
                  {taskState[r.id] === "待触达" ? "记录触达" : "完成并回流"}
                </Button>
              ),
            },
          ]}
        />
        <div className="feedback-loop">
          <span>
            <b>回流内容</b>触达结果、融资需求、授信受理、审批结论、风险处置
          </span>
          <em>→</em>
          <span>
            <b>自动更新</b>线索转化、授信转化、预警提前量、误报率
          </span>
          <em>→</em>
          <span>
            <b>建设端迭代</b>效果归因、阈值调整、模型版本升级
          </span>
        </div>
      </Card>
    </div>
  );
}

function ProductDetail({
  published,
  go,
}: {
  published: boolean;
  go: (v: View) => void;
}) {
  const selectedName =
    typeof window !== "undefined"
      ? localStorage.getItem("icbc-selected-product")
      : null;
  const publishedRecord =
    typeof window !== "undefined"
      ? JSON.parse(localStorage.getItem("icbc-published-product") || "null")
      : null;
  const name =
    selectedName ||
    (published
      ? publishedRecord?.name || "全国产业链机会与风险识别"
      : "全国产业链机会与风险识别");
  const catalogProduct = initialProducts.find((p: any) => p.name === name);
  const isPublishedEvent =
    published && name === (publishedRecord?.name || "全国产业链机会与风险识别");
  const detailProfile =
    catalogProduct ||
    (isPublishedEvent
      ? {
          category: "客户营销＋风险监测",
          scope: "全行",
          desc: "以产业事件为触发，融合企业基本面、外部信用与行内经营信息，持续更新机会、风险和客户关系判断。",
          users: "对公客户经理、风险管理人员",
          uses: 0,
          rating: 5,
          status: "新发布",
        }
      : {
          category: "客户营销",
          scope: "全行",
          desc: "筛选符合地区、行业、资质、经营活跃度与风险要求的企业，形成可解释的业务结果并推送后续行动。",
          users: "对公客户经理",
          uses: 0,
          rating: 5,
          status: "新发布",
        });
  const canonicalProduct = "全国产业链机会与风险识别";
  const runnableConfig =
    workbenchProducts[name] ||
    (isPublishedEvent ? workbenchProducts[canonicalProduct] : null);
  const runnable = !!runnableConfig;
  const embeddedRelation = name.includes("司法")
    ? { title: "风险事件关联", scene: "风险监测", resources: ["司法涉诉查询", "经营异常与失信", "社会舆情资讯"], abilities: ["主体关联", "司法风险核验", "舆情事件识别"], rules: ["司法风险分级", "经营异常排除"] }
    : name.includes("产业链") || name.includes("产业")
      ? { title: "产业链关系图谱", scene: "产业研究与经营决策", resources: ["产业链企业图谱", "工商登记数据库", "产业事件资讯"], abilities: ["主体关联", "产业链节点识别", "影响传导分析"], rules: ["地区参数", "行业参数", "事件窗口"] }
      : { title: "营销线索关联", scene: "客户拓展", resources: ["工商登记数据库", "企业资质名单", "招投标公告库"], abilities: ["主体关联", "产业资质识别", "经营活跃度计算"], rules: ["地区筛选", "资质筛选", "风险排除"] };
  const detailResources = runnable
    ? resources.filter((r) => runnableConfig.resources.includes(r.id))
    : resources.slice(0, 4);
  const taskCopy =
    name === canonicalProduct
      ? {
          problem:
            "回答产业事件如何沿产业链传导、影响哪些链上客户、形成何种融资机会与风险，并将判断转化为跨岗位协同任务。",
          flow: [
            "产业链范围设定",
            "多源事件触发",
            "影响传导分析",
            "客户机会与风险决策",
            "协同任务与反馈",
          ],
          role: "对公客户经理、风险管理人员、经营分析人员",
          output:
            "产业链经营驾驶舱、事件影响图、客户分层、融资机会包、风险传导路径与行动建议",
          action: "进入营销、授信、风险或人工复核任务",
        }
      : detailProfile.category === "风险监测"
        ? {
            problem:
              "将多源风险事件统一关联至客户主体，形成分级预警、影响分析和处置任务。",
            flow: [
              "客户范围确定",
              "风险事件归集",
              "风险分级",
              "影响核验",
              "预警推送",
            ],
            role: "风险管理人员",
            output: "风险事件、预警等级、处置建议",
            action: "进入风险处置任务池",
          }
        : detailProfile.category === "授信支持"
          ? {
              problem:
                "整合工商、股权、司法、处罚和财务信息，减少授信调查中的跨平台查询与人工核验。",
              flow: [
                "客户输入",
                "主体核验",
                "关系穿透",
                "风险核验",
                "尽调报告",
              ],
              role: "客户经理、授信审查人员",
              output: "外部信息核验清单、风险摘要",
              action: "进入授信调查流程",
            }
          : detailProfile.category === "产业研究" ||
              detailProfile.category === "经营分析"
            ? {
                problem:
                  "将产业链、企业经营和区域景气数据组合分析，形成结构化画像与趋势判断。",
                flow: [
                  "研究范围设定",
                  "数据归集",
                  "指标计算",
                  "结构分析",
                  "结果交付",
                ],
                role: "产业研究、经营分析人员",
                output: "产业画像、指标看板、企业清单",
                action: "用于研究分析与经营决策",
              }
            : detailProfile.category === "监管与合规"
              ? {
                  problem:
                    "持续归集行政处罚、监管名单和合规事件，形成统一筛查结果与跟踪提醒。",
                  flow: [
                    "主体输入",
                    "名单筛查",
                    "事项归集",
                    "合规分级",
                    "提醒处置",
                  ],
                  role: "合规管理、国际业务人员",
                  output: "合规事项、命中依据、处置提醒",
                  action: "进入合规核验与处置流程",
                }
              : {
                  problem:
                    "将企业资质、经营活动和风险信息转化为可解释的营销线索和目标客户名单。",
                  flow: [
                    "目标客户发现",
                    "企业初筛",
                    "名单形成",
                    "客户触达",
                    "营销反馈",
                  ],
                  role: detailProfile.users || "对公客户经理、营销主管",
                  output: "营销名单、企业画像、推荐理由",
                  action: "进入客户经理任务池",
                };
  const [followed, setFollowed] = useState(() => {
    const raw =
      typeof window === "undefined"
        ? null
        : localStorage.getItem("icbc-follow-products");
    const stored = raw ? JSON.parse(raw) : defaultFollowProducts;
    return stored.includes(name);
  });
  const toggleFollow = () => {
    const raw = localStorage.getItem("icbc-follow-products");
    const old = raw ? JSON.parse(raw) : defaultFollowProducts;
    const next = old.includes(name)
      ? old.filter((x: string) => x !== name)
      : [...old, name];
    localStorage.setItem("icbc-follow-products", JSON.stringify(next));
    setFollowed(next.includes(name));
    message.success(next.includes(name) ? "已加入我的关注" : "已取消关注");
  };
  const openWorkbench = () => {
    if (runnable) {
      localStorage.setItem("icbc-workbench-product", name);
      go("workbench");
    } else message.info("该产品当前提供目录与详情演示，运行配置正在补充");
  };
  return (
    <div>
      <Breadcrumb
        items={[
          { title: "外数产品" },
          { title: detailProfile.category },
          { title: name },
        ]}
      />
      <div className="product-detail-head">
        <div>
          <Space>
            <Tag color="red">{detailProfile.category}</Tag>
            <Tag>{detailProfile.scope}</Tag>
            <StatusTag status={detailProfile.status || "正常运行"} />
          </Space>
          <h1>{name}</h1>
          <p>{detailProfile.desc}</p>
          <Space>
            <Button type="primary" size="large" onClick={openWorkbench}>
              {runnable ? "立即使用" : "查看产品目录"}
            </Button>
            <Button
              size="large"
              type={followed ? "primary" : "default"}
              icon={followed ? <CheckCircleFilled /> : <HeartOutlined />}
              onClick={toggleFollow}
            >
              {followed ? "已关注" : "关注产品"}
            </Button>
          </Space>
        </div>
        <div className="detail-stats">
          <Statistic title="当前版本" value="V1.0" />
          <Statistic title="累计使用" value={detailProfile.uses || 0} />
          <Statistic title="用户评分" value={detailProfile.rating || 5} />
          <Statistic title="结果采纳率" value={78} suffix="%" />
        </div>
      </div>
      <Card>
        <Tabs
          items={[
            {
              key: "overview",
              label: "产品概览",
              children: (
                <Row gutter={20}>
                  <Col span={15}>
                    <SectionTitle title="产品解决什么问题" />
                    <p className="large-copy">{taskCopy.problem}</p>
                    <SectionTitle title="业务流程位置" />
                    <div className="business-flow small">
                      {taskCopy.flow.map((x: string, i: number) => (
                        <React.Fragment key={x}>
                          <button className={i === 1 ? "selected" : ""}>
                            <span>{i + 1}</span>
                            <b>{x}</b>
                          </button>
                          {i < taskCopy.flow.length - 1 && <em>→</em>}
                        </React.Fragment>
                      ))}
                    </div>
                  </Col>
                  <Col span={9}>
                    <SectionTitle title="适用岗位与结果" />
                    <Descriptions
                      column={1}
                      bordered
                      size="small"
                      items={[
                        { label: "主要岗位", children: taskCopy.role },
                        {
                          label: "输入条件",
                          children:
                            "地域、产业链、事件窗口、客户范围与业务阈值",
                        },
                        { label: "输出结果", children: taskCopy.output },
                        { label: "后续动作", children: taskCopy.action },
                      ]}
                    />
                  </Col>
                </Row>
              ),
            },
            ...(name === canonicalProduct
              ? [
                  {
                    key: "command",
                    label: "决策驾驶舱预览",
                    children: <IndustryCommandPreview />,
                  },
                ]
              : []),
            {
              key: "relation",
              label: embeddedRelation.title,
              children: (
                <Row gutter={16}>
                  <Col span={17}>
                    <ReactECharts style={{ height: 410 }} option={{
                      tooltip: {},
                      series: [{ type: "graph", layout: "force", roam: true, label: { show: true, fontSize: 11 }, force: { repulsion: 300, edgeLength: 110 },
                        data: [
                          { name, symbolSize: 66, itemStyle: { color: "#c7000b" } },
                          { name: embeddedRelation.scene, symbolSize: 46, itemStyle: { color: "#316bad" } },
                          ...embeddedRelation.resources.map((x: string) => ({ name: x, symbolSize: 38, itemStyle: { color: "#28966f" } })),
                          ...embeddedRelation.abilities.map((x: string) => ({ name: x, symbolSize: 38, itemStyle: { color: "#d77c00" } })),
                          ...embeddedRelation.rules.map((x: string) => ({ name: x, symbolSize: 32, itemStyle: { color: "#7651a8" } })),
                        ],
                        links: [embeddedRelation.scene, ...embeddedRelation.resources, ...embeddedRelation.abilities, ...embeddedRelation.rules].map((x: string) => ({ source: name, target: x }))
                      }]
                    }} />
                  </Col>
                  <Col span={7}>
                    <Card size="small" title="关联摘要">
                      <Statistic title="关联外数资源" value={embeddedRelation.resources.length} suffix="项" />
                      <Statistic title="复用能力" value={embeddedRelation.abilities.length} suffix="项" />
                      <Statistic title="调用规则" value={embeddedRelation.rules.length} suffix="条" />
                      <p className="muted">资源更新、能力升级和规则调整均可追溯至本产品版本。</p>
                    </Card>
                  </Col>
                </Row>
              ),
            },
            {
              key: "process",
              label: "加工过程",
              children: (
                <div className="process-business">
                  {[
                    "主体关联",
                    "指标计算",
                    "规则筛选",
                    "风险核验",
                    "名单生成",
                    "行动提示",
                  ].map((x, i) => (
                    <React.Fragment key={x}>
                      <div>
                        <span>{i + 1}</span>
                        <b>{x}</b>
                        <small>
                          {
                            [
                              "统一识别企业",
                              "形成活跃度等指标",
                              "应用业务条件",
                              "排除重大风险",
                              "按优先级排序",
                              "说明推荐原因",
                            ][i]
                          }
                        </small>
                      </div>
                      {i < 5 && <em>→</em>}
                    </React.Fragment>
                  ))}
                </div>
              ),
            },
            {
              key: "source",
              label: `数据来源（${detailResources.length}）`,
              children: (
                <Table
                  pagination={false}
                  dataSource={detailResources}
                  rowKey="id"
                  columns={[
                    { title: "资源名称", dataIndex: "name" },
                    { title: "资源类型", dataIndex: "type" },
                    { title: "更新频率", dataIndex: "frequency" },
                    {
                      title: "授权状态",
                      dataIndex: "auth",
                      render: (x: string) => <StatusTag status={x} />,
                    },
                    { title: "支撑环节", render: () => "数据输入与业务核验" },
                  ]}
                />
              ),
            },
            ...(name === canonicalProduct
              ? [
                  {
                    key: "collaboration",
                    label: "与 e企查协同",
                    children: (
                      <div className="equery-collab">
                        <div>
                          <DatabaseOutlined />
                          <span>
                            <b>外数产品管理平台</b>
                            <small>
                              生产和治理动态评估：资源、能力、规则、版本与运行
                            </small>
                          </span>
                        </div>
                        <em>输出评估服务</em>
                        <div>
                          <SearchOutlined />
                          <span>
                            <b>e企查</b>
                            <small>
                              企业查询入口：展示评估卡片、变化原因与行动入口
                            </small>
                          </span>
                        </div>
                        <Card size="small">
                          <b>平台边界</b>
                          <p>
                            e企查回答企业是谁、有哪些工商司法和关联信息；本平台回答如何把内外部数据加工成产品、形成何种判断并持续运营。动态企业评估是平台生产的一项代表性产品，成熟后可嵌入
                            e企查 丰富企业详情页。
                          </p>
                          <Button
                            onClick={() =>
                              message.success("已生成 e企查 动态评估卡片预览")
                            }
                          >
                            预览嵌入卡片
                          </Button>
                        </Card>
                      </div>
                    ),
                  },
                ]
              : []),
            {
              key: "version",
              label: "版本与评价",
              children: (
                <Timeline
                  items={
                    name === canonicalProduct
                      ? [
                          {
                            children:
                              "V3.1 · 新增三维动态评估、变化归因与内外数融合",
                          },
                          { children: "V3.0 · 产业事件机会与风险双向识别" },
                          { children: "测试报告 TEST-2026-0805-12 全部通过" },
                        ]
                      : [
                          { children: "V1.0 · 2026-07-31 首次发布" },
                          { children: "测试报告全部通过" },
                        ]
                  }
                />
              ),
            },
          ]}
        />
      </Card>
    </div>
  );
}

const workbenchProducts: Record<string, any> = {
  全国产业链机会与风险识别: {
    resources: [
      "EXT-R001",
      "EXT-R002",
      "EXT-R003",
      "EXT-R004",
      "EXT-R005",
      "EXT-R007",
      "EXT-R008",
      "EXT-R015",
      "EXT-R019",
      "EXT-R020",
      "EXT-R022",
    ],
    abilities: [
      "主体关联",
      "产业链节点识别",
      "事件抽取与标准化",
      "机会评分",
      "风险核验",
      "影响传导分析",
      "策略推荐与解释",
    ],
    conditions: [
      {
        key: "region",
        label: "地区",
        values: [
          "全国",
          "四川省",
          "北京市",
          "上海市",
          "江苏省",
          "浙江省",
          "广东省",
        ],
      },
      {
        key: "industry",
        label: "行业或产业链",
        values: [
          "动力电池",
          "新能源汽车",
          "高端装备",
          "集成电路",
          "光伏",
          "储能",
          "机器人",
          "生物医药",
        ],
      },
      {
        key: "chain",
        label: "产业链环节",
        values: ["全产业链", "上游材料", "核心部件", "整机/系统", "回收与服务"],
      },
      {
        key: "event",
        label: "事件类型",
        values: [
          "全部机会与风险事件",
          "重大项目与扩产",
          "招中标",
          "融资与投资",
          "政策与舆情",
          "司法与经营风险",
        ],
      },
      {
        key: "window",
        label: "观察期限",
        values: ["近30日", "近90日", "近半年", "近一年"],
      },
      {
        key: "company",
        label: "企业类型",
        values: [
          "全部企业",
          "存量客户",
          "潜在客户",
          "专精特新企业",
          "国有企业",
        ],
      },
      {
        key: "opportunity",
        label: "机会阈值",
        values: ["70分", "80分", "90分"],
      },
      {
        key: "risk",
        label: "风险阈值",
        values: ["不高于60分", "不高于40分", "仅排除重大风险"],
      },
      {
        key: "limit",
        label: "推荐企业数量",
        values: ["20家", "50家", "100家"],
      },
    ],
  },
  浙江分行专精特新企业筛选产品: {
    resources: ["EXT-R001", "EXT-R003", "EXT-R004", "EXT-R005"],
    abilities: ["主体关联", "产业资质识别", "招投标活跃度计算", "司法风险核验"],
    conditions: [
      {
        key: "region",
        label: "地区",
        values: ["浙江省", "杭州市", "宁波市", "嘉兴市"],
      },
      {
        key: "industry",
        label: "行业",
        values: ["制造业", "高端装备", "新材料", "工业软件"],
      },
      { key: "years", label: "注册年限", values: ["近三年", "近五年", "不限"] },
      {
        key: "qualification",
        label: "企业资质",
        values: ["专精特新小巨人", "省级专精特新"],
      },
      {
        key: "bid",
        label: "招投标记录",
        values: ["近一年存在", "近半年存在", "不限"],
      },
      {
        key: "risk",
        label: "司法风险",
        values: ["近三个月无重大风险", "近六个月无重大风险"],
      },
    ],
  },
  全行专精特新企业营销名单: {
    resources: ["EXT-R001", "EXT-R003", "EXT-R004", "EXT-R005", "EXT-R006"],
    abilities: [
      "主体关联",
      "工商信息标准化",
      "产业资质识别",
      "名单筛选",
      "司法风险核验",
    ],
    conditions: [
      {
        key: "region",
        label: "地区",
        values: ["全国", "北京市", "上海市", "江苏省", "浙江省", "广东省"],
      },
      {
        key: "industry",
        label: "行业",
        values: ["制造业", "信息技术", "高端装备", "新材料"],
      },
      { key: "years", label: "注册年限", values: ["近三年", "近五年", "不限"] },
      {
        key: "qualification",
        label: "企业资质",
        values: ["国家级小巨人", "省级专精特新"],
      },
      {
        key: "scale",
        label: "企业规模",
        values: ["不限", "中型企业", "小微企业"],
      },
      {
        key: "risk",
        label: "风险条件",
        values: ["无重大司法及经营风险", "仅排除高风险"],
      },
    ],
  },
  企业司法风险监测: {
    resources: [
      "EXT-R001",
      "EXT-R005",
      "EXT-R006",
      "EXT-R007",
      "EXT-R025",
      "EXT-R026",
    ],
    abilities: ["主体关联", "司法风险核验", "舆情事件识别", "风险分级"],
    conditions: [
      {
        key: "region",
        label: "监测地域",
        values: ["全国", "北京市", "上海市", "广东省", "浙江省"],
      },
      {
        key: "customer",
        label: "客户范围",
        values: ["全部存量客户", "重点客户", "授信客户"],
      },
      {
        key: "window",
        label: "监测窗口",
        values: ["近三个月", "近六个月", "近一年"],
      },
      {
        key: "event",
        label: "事件类型",
        values: ["全部风险事件", "被执行与失信", "重大诉讼", "经营异常"],
      },
      { key: "level", label: "预警等级", values: ["中风险及以上", "仅高风险"] },
      { key: "frequency", label: "运行频率", values: ["每日", "每周"] },
    ],
  },
  招投标客户发现: {
    resources: ["EXT-R001", "EXT-R003", "EXT-R005", "EXT-R013"],
    abilities: ["主体关联", "招投标活跃度计算", "融资机会评分", "名单筛选"],
    conditions: [
      {
        key: "region",
        label: "项目地区",
        values: ["全国", "北京市", "江苏省", "浙江省", "广东省"],
      },
      {
        key: "industry",
        label: "所属行业",
        values: ["制造业", "建筑业", "信息技术", "新能源"],
      },
      {
        key: "window",
        label: "公告窗口",
        values: ["近三个月", "近半年", "近一年"],
      },
      {
        key: "amount",
        label: "项目金额",
        values: ["1000万元以上", "500万元以上", "不限"],
      },
      {
        key: "notice",
        label: "公告类型",
        values: ["中标公告", "招标与中标公告", "全部"],
      },
      { key: "risk", label: "风险条件", values: ["排除高风险企业", "不限"] },
    ],
  },
  授信尽调外部信息核验: {
    resources: [
      "EXT-R001",
      "EXT-R002",
      "EXT-R005",
      "EXT-R006",
      "EXT-R021",
      "EXT-R030",
      "EXT-R037",
    ],
    abilities: ["主体关联", "股权穿透", "司法风险核验", "外部信息核验"],
    conditions: [
      {
        key: "region",
        label: "客户地域",
        values: ["全国", "北京市", "上海市", "江苏省", "广东省"],
      },
      {
        key: "customer",
        label: "客户类型",
        values: ["拟授信客户", "存量授信客户", "集团客户"],
      },
      {
        key: "depth",
        label: "股权穿透层级",
        values: ["三层", "五层", "穿透至最终受益人"],
      },
      {
        key: "window",
        label: "风险回溯期",
        values: ["近一年", "近三年", "全部有效记录"],
      },
      {
        key: "material",
        label: "核验范围",
        values: ["工商、股权、司法、处罚", "全部外部信息"],
      },
      { key: "frequency", label: "报告更新", values: ["按次生成", "每日监测"] },
    ],
  },
  舆情风险预警: {
    resources: ["EXT-R001", "EXT-R007", "EXT-R018", "EXT-R021"],
    abilities: ["主体关联", "舆情事件识别", "情感分析", "风险分级"],
    conditions: [
      {
        key: "region",
        label: "监测地域",
        values: ["全国", "重点区域", "指定机构"],
      },
      {
        key: "customer",
        label: "监测客群",
        values: ["全部存量客户", "重点授信客户", "自定义名单"],
      },
      {
        key: "event",
        label: "事件主题",
        values: ["全部负面事件", "经营异常", "安全事故", "治理与声誉"],
      },
      {
        key: "source",
        label: "信息来源",
        values: ["权威媒体优先", "全网公开信息"],
      },
      { key: "level", label: "预警阈值", values: ["中风险及以上", "仅高风险"] },
      { key: "frequency", label: "推送频率", values: ["实时", "每日汇总"] },
    ],
  },
  区域产业链客户画像: {
    resources: [
      "EXT-R001",
      "EXT-R008",
      "EXT-R011",
      "EXT-R019",
      "EXT-R022",
      "EXT-R040",
    ],
    abilities: ["主体关联", "产业链节点识别", "企业画像", "关系图谱分析"],
    conditions: [
      {
        key: "region",
        label: "研究区域",
        values: ["长三角", "京津冀", "粤港澳大湾区", "成渝地区"],
      },
      {
        key: "chain",
        label: "产业链",
        values: ["新能源汽车", "集成电路", "高端装备", "生物医药"],
      },
      {
        key: "role",
        label: "链属角色",
        values: ["全部企业", "链主企业", "核心配套企业"],
      },
      { key: "depth", label: "关系层级", values: ["上下游一层", "上下游三层"] },
      {
        key: "indicator",
        label: "画像重点",
        values: ["经营与融资", "技术与资质", "风险与依赖度"],
      },
      { key: "period", label: "观察期", values: ["近一年", "近三年"] },
    ],
  },
  行政处罚与合规事项监测: {
    resources: ["EXT-R001", "EXT-R021", "EXT-R028", "EXT-R029", "EXT-R039"],
    abilities: ["主体关联", "监管事项归集", "合规规则匹配", "风险分级"],
    conditions: [
      {
        key: "region",
        label: "覆盖地域",
        values: ["全国", "北京市", "上海市", "广东省"],
      },
      {
        key: "customer",
        label: "客户范围",
        values: ["全部客户", "授信客户", "跨境业务客户"],
      },
      {
        key: "event",
        label: "事项类型",
        values: ["全部合规事项", "行政处罚", "环保处罚", "税务风险"],
      },
      {
        key: "window",
        label: "监测窗口",
        values: ["近三个月", "近一年", "持续监测"],
      },
      { key: "level", label: "提醒等级", values: ["一般及以上", "重大事项"] },
      { key: "frequency", label: "更新频率", values: ["每日", "每周"] },
    ],
  },
};

const judicialResults = [
  [
    "华东智造装备集团有限公司",
    "上海市",
    "被执行",
    "(2026)沪01执1842号",
    2860000,
    "高",
    "2026-07-29",
  ],
  [
    "粤海供应链科技股份有限公司",
    "广州市",
    "重大诉讼",
    "(2026)粤01民初736号",
    12500000,
    "高",
    "2026-07-27",
  ],
  [
    "北辰精密制造有限公司",
    "北京市",
    "经营异常",
    "京市监异列字〔2026〕418号",
    0,
    "中",
    "2026-07-25",
  ],
  [
    "浙能工业软件有限公司",
    "杭州市",
    "裁判文书",
    "(2026)浙01民终2291号",
    780000,
    "中",
    "2026-07-22",
  ],
  [
    "苏南新材料股份有限公司",
    "苏州市",
    "开庭公告",
    "(2026)苏05民初905号",
    460000,
    "中",
    "2026-07-19",
  ],
  [
    "鲁华新能源设备有限公司",
    "济南市",
    "失信被执行",
    "(2026)鲁01执恢318号",
    1960000,
    "高",
    "2026-07-17",
  ],
  [
    "闽锐电子材料有限公司",
    "厦门市",
    "财产保全",
    "(2026)闽02财保156号",
    1320000,
    "中",
    "2026-07-15",
  ],
  [
    "中原智能农机股份有限公司",
    "郑州市",
    "合同纠纷",
    "(2026)豫01民初1248号",
    980000,
    "中",
    "2026-07-13",
  ],
  [
    "湘江生物医药有限公司",
    "长沙市",
    "开庭公告",
    "(2026)湘01民初771号",
    350000,
    "中",
    "2026-07-11",
  ],
  [
    "渝西精密传动有限公司",
    "重庆市",
    "被执行",
    "(2026)渝05执906号",
    4380000,
    "高",
    "2026-07-09",
  ],
  [
    "津港绿色物流有限公司",
    "天津市",
    "经营异常",
    "津市监异列字〔2026〕284号",
    0,
    "中",
    "2026-07-07",
  ],
  [
    "陕科储能系统有限公司",
    "西安市",
    "裁判文书",
    "(2026)陕01民终1836号",
    620000,
    "中",
    "2026-07-05",
  ],
].map((x, i) => ({
  key: i + 1,
  name: x[0],
  city: x[1],
  event: x[2],
  caseNo: x[3],
  amount: x[4],
  risk: x[5],
  date: x[6],
}));

const workbenchRuntime: Record<string, any> = {
  全国产业链机会与风险识别: {
    unit: "家重点企业",
    empty: "配置地区、产业链、事件、期限和阈值后生成综合研判",
    success: "全国产业链综合研判完成",
    push: "按建议方向推送营销、授信、风险或复核任务",
    file: "全国产业链机会与风险识别企业明细_演示.xls",
    fields: [
      "企业名称:name",
      "触发事件:event",
      "事件摘要:summary",
      "事件金额:amount",
      "机会评分:score",
      "风险等级:risk",
      "建议业务方向:action",
      "事件日期:date",
    ],
    rows: [
      {
        name: "蜀能智能装备有限公司",
        event: "重大项目中标",
        summary: "中标新能源装备扩产项目",
        amount: 12800,
        score: 94,
        risk: "低",
        action: "联系项目融资需求",
        date: "2026-08-02",
      },
      {
        name: "成都新材科技有限公司",
        event: "扩产技改",
        summary: "公告建设高性能材料二期产线",
        amount: 7600,
        score: 89,
        risk: "低",
        action: "提供技改贷款方案",
        date: "2026-08-01",
      },
      {
        name: "川南化工集团",
        event: "行政处罚",
        summary: "生产安全事项被责令整改",
        amount: 260,
        score: 42,
        risk: "高",
        action: "推送风险人员核验",
        date: "2026-07-31",
      },
      {
        name: "西部算力科技有限公司",
        event: "政策支持",
        summary: "入选省级算力基础设施重点项目",
        amount: 9800,
        score: 91,
        risk: "低",
        action: "跟进项目贷款与设备融资",
        date: "2026-07-29",
      },
      {
        name: "德阳航空零部件有限公司",
        event: "战略合作",
        summary: "与链主企业签署三年核心部件供货协议",
        amount: 6200,
        score: 87,
        risk: "低",
        action: "核实订单融资及保函需求",
        date: "2026-07-28",
      },
      {
        name: "绵阳功率半导体科技有限公司",
        event: "产能扩建",
        summary: "新增车规级功率器件封测产线",
        amount: 15600,
        score: 93,
        risk: "低",
        action: "组建技改融资联合服务方案",
        date: "2026-07-27",
      },
      {
        name: "宜宾动力材料股份有限公司",
        event: "再融资需求",
        summary: "发布中期票据注册计划",
        amount: 30000,
        score: 90,
        risk: "中低",
        action: "联动投行团队跟进承销机会",
        date: "2026-07-25",
      },
      {
        name: "泸州绿色化工有限公司",
        event: "环保处罚",
        summary: "废气排放指标超限并启动限期整改",
        amount: 180,
        score: 38,
        risk: "高",
        action: "暂停营销并发起授信风险复核",
        date: "2026-07-23",
      },
      {
        name: "眉山智能仓储有限公司",
        event: "项目备案",
        summary: "智慧物流园项目完成备案",
        amount: 11800,
        score: 86,
        risk: "低",
        action: "跟进项目贷与供应链融资",
        date: "2026-07-21",
      },
      {
        name: "乐山新型储能有限公司",
        event: "政府补助",
        summary: "获得省级新型储能示范项目专项补助",
        amount: 2400,
        score: 88,
        risk: "低",
        action: "匹配绿色金融专项产品",
        date: "2026-07-19",
      },
      {
        name: "遂宁电子信息集团",
        event: "高风险舆情",
        summary: "主要客户集中度上升引发经营压力关注",
        amount: 0,
        score: 51,
        risk: "中",
        action: "核验订单稳定性并调整营销优先级",
        date: "2026-07-17",
      },
      {
        name: "达州先进材料有限公司",
        event: "重大诉讼",
        summary: "涉及原材料采购合同争议",
        amount: 4600,
        score: 35,
        risk: "高",
        action: "推送风险人员核验敞口影响",
        date: "2026-07-15",
      },
    ],
  },
  浙江分行专精特新企业筛选产品: {
    unit: "家企业",
    empty: "配置地域、资质与活跃度条件后生成浙江营销名单",
    success: "筛选完成",
    push: "推送浙江客户经理任务池",
    file: "浙江专精特新企业筛选结果_演示.xls",
    fields: [
      "企业名称:name",
      "地区:city",
      "所属行业:industry",
      "企业资质:qualification",
      "招投标次数:bids",
      "推荐分:score",
      "风险状态:risk",
    ],
    rows: companyResults,
  },
  全行专精特新企业营销名单: {
    unit: "家企业",
    empty: "配置客群条件后生成全行专精特新营销名单",
    success: "筛选完成",
    push: "推送客户经理任务池",
    file: "全行专精特新营销名单_演示.xls",
    fields: [
      "企业名称:name",
      "地区:city",
      "所属行业:industry",
      "企业资质:qualification",
      "招投标次数:bids",
      "推荐分:score",
      "风险状态:risk",
    ],
    rows: companyResults,
  },
  企业司法风险监测: {
    unit: "条风险事件",
    empty: "配置客户范围与风险阈值后运行监测",
    success: "风险监测完成",
    push: "推送风险处置任务池",
    file: "企业司法风险监测结果_演示.xls",
    fields: [
      "企业名称:name",
      "地区:city",
      "事件类型:event",
      "案号／文号:caseNo",
      "涉案金额:amount",
      "风险等级:risk",
      "事件日期:date",
    ],
    rows: judicialResults,
  },
  招投标客户发现: {
    unit: "条融资线索",
    empty: "配置项目地域、金额与公告窗口后识别融资线索",
    success: "招投标商机识别完成",
    push: "推送营销商机池",
    file: "招投标客户发现结果_演示.xls",
    fields: [
      "企业名称:name",
      "项目地区:city",
      "中标项目:project",
      "中标金额:amount",
      "公告日期:date",
      "融资机会分:score",
      "风险状态:risk",
    ],
    rows: [
      {
        name: "川西智造装备有限公司",
        city: "成都市",
        project: "智能产线升级项目",
        amount: 8600,
        date: "2026-08-01",
        score: 92,
        risk: "低",
      },
      {
        name: "苏南储能科技有限公司",
        city: "苏州市",
        project: "园区储能系统采购",
        amount: 5200,
        date: "2026-07-30",
        score: 88,
        risk: "低",
      },
      {
        name: "粤港冷链物流有限公司",
        city: "广州市",
        project: "智慧冷链基地建设",
        amount: 4100,
        date: "2026-07-28",
        score: 84,
        risk: "中低",
      },
      {
        name: "京北智慧交通有限公司",
        city: "北京市",
        project: "城市交通信号智能化改造",
        amount: 7300,
        date: "2026-07-26",
        score: 90,
        risk: "低",
      },
      {
        name: "甬海高端装备有限公司",
        city: "宁波市",
        project: "港口自动化装卸设备采购",
        amount: 6800,
        date: "2026-07-25",
        score: 89,
        risk: "低",
      },
      {
        name: "皖中新材料股份有限公司",
        city: "合肥市",
        project: "新能源电池材料供应项目",
        amount: 4600,
        date: "2026-07-23",
        score: 85,
        risk: "低",
      },
      {
        name: "鲁能环保工程有限公司",
        city: "济南市",
        project: "工业园区污水处理扩容工程",
        amount: 3900,
        date: "2026-07-21",
        score: 82,
        risk: "中低",
      },
      {
        name: "闽创信息系统有限公司",
        city: "福州市",
        project: "政务云安全能力升级项目",
        amount: 2800,
        date: "2026-07-19",
        score: 81,
        risk: "低",
      },
      {
        name: "鄂州医疗设备有限公司",
        city: "武汉市",
        project: "区域医学中心设备采购",
        amount: 6100,
        date: "2026-07-18",
        score: 87,
        risk: "低",
      },
      {
        name: "西安精密测控有限公司",
        city: "西安市",
        project: "工业检测平台建设项目",
        amount: 3500,
        date: "2026-07-16",
        score: 83,
        risk: "中低",
      },
    ],
  },
  授信尽调外部信息核验: {
    unit: "家待核验客户",
    empty: "输入拟授信客户并配置穿透层级后生成核验报告",
    success: "外部信息核验完成",
    push: "推送授信调查流程",
    file: "授信尽调外部核验结果_演示.xls",
    fields: [
      "企业名称:name",
      "注册状态:status",
      "最终受益人:ubo",
      "关联企业数:relations",
      "司法事项:judicial",
      "行政处罚:penalty",
      "核验结论:conclusion",
    ],
    rows: [
      {
        name: "华东精工集团有限公司",
        status: "存续",
        ubo: "周某",
        relations: 12,
        judicial: "2项一般诉讼",
        penalty: "无",
        conclusion: "建议通过",
      },
      {
        name: "北辰新材股份有限公司",
        status: "存续",
        ubo: "李某",
        relations: 7,
        judicial: "无",
        penalty: "1项已整改",
        conclusion: "人工复核",
      },
      {
        name: "粤科电子有限公司",
        status: "存续",
        ubo: "陈某",
        relations: 18,
        judicial: "1项被执行",
        penalty: "无",
        conclusion: "重点核验",
      },
      {
        name: "苏北新能源装备有限公司",
        status: "存续",
        ubo: "王某",
        relations: 9,
        judicial: "无",
        penalty: "无",
        conclusion: "建议通过",
      },
      {
        name: "泉州精密制造有限公司",
        status: "存续",
        ubo: "林某",
        relations: 15,
        judicial: "1项合同纠纷",
        penalty: "无",
        conclusion: "建议通过",
      },
      {
        name: "长沙数智工程股份有限公司",
        status: "存续",
        ubo: "赵某",
        relations: 21,
        judicial: "无",
        penalty: "2项已整改",
        conclusion: "人工复核",
      },
      {
        name: "重庆新材科技集团",
        status: "存续",
        ubo: "刘某",
        relations: 26,
        judicial: "3项一般诉讼",
        penalty: "1项整改中",
        conclusion: "重点核验",
      },
      {
        name: "河北绿色建材有限公司",
        status: "存续",
        ubo: "孙某",
        relations: 8,
        judicial: "无",
        penalty: "无",
        conclusion: "建议通过",
      },
      {
        name: "青岛海洋装备有限公司",
        status: "存续",
        ubo: "高某",
        relations: 13,
        judicial: "1项开庭公告",
        penalty: "无",
        conclusion: "建议通过",
      },
      {
        name: "云南生物医药有限公司",
        status: "存续",
        ubo: "杨某",
        relations: 11,
        judicial: "无",
        penalty: "1项轻微已整改",
        conclusion: "建议通过",
      },
    ],
  },
  舆情风险预警: {
    unit: "条舆情事件",
    empty: "配置客群、主题与预警阈值后运行舆情监测",
    success: "舆情监测完成",
    push: "推送声誉风险处置",
    file: "舆情风险预警结果_演示.xls",
    fields: [
      "企业名称:name",
      "事件主题:event",
      "信息来源:source",
      "事件摘要:summary",
      "情感倾向:sentiment",
      "预警等级:risk",
      "发布时间:date",
    ],
    rows: [
      {
        name: "东江化工有限公司",
        event: "安全生产",
        source: "省应急管理厅",
        summary: "生产装置停产整改",
        sentiment: "负面",
        risk: "高",
        date: "2026-08-02",
      },
      {
        name: "北方零部件集团",
        event: "经营波动",
        source: "权威财经媒体",
        summary: "主要客户订单调整",
        sentiment: "负面",
        risk: "中",
        date: "2026-08-01",
      },
      {
        name: "西南数字科技有限公司",
        event: "公司治理",
        source: "交易所公告",
        summary: "董事长发生变更",
        sentiment: "中性",
        risk: "中",
        date: "2026-07-31",
      },
      {
        name: "江海新能源股份有限公司",
        event: "产品质量",
        source: "市场监管部门",
        summary: "部分批次产品启动主动召回",
        sentiment: "负面",
        risk: "高",
        date: "2026-07-30",
      },
      {
        name: "中州物流集团",
        event: "经营波动",
        source: "行业媒体",
        summary: "运价下降导致季度盈利承压",
        sentiment: "负面",
        risk: "中",
        date: "2026-07-28",
      },
      {
        name: "宁波精密传动有限公司",
        event: "知识产权",
        source: "法院公告",
        summary: "核心专利纠纷进入审理程序",
        sentiment: "负面",
        risk: "中",
        date: "2026-07-27",
      },
      {
        name: "鹏城数字服务有限公司",
        event: "数据安全",
        source: "监管通报",
        summary: "因个人信息保护问题被要求整改",
        sentiment: "负面",
        risk: "高",
        date: "2026-07-25",
      },
      {
        name: "合肥新型显示有限公司",
        event: "项目进展",
        source: "地方政府网站",
        summary: "新产线按期完成设备进场",
        sentiment: "正面",
        risk: "低",
        date: "2026-07-23",
      },
      {
        name: "津门食品科技有限公司",
        event: "供应链",
        source: "权威财经媒体",
        summary: "主要原料价格短期明显上涨",
        sentiment: "负面",
        risk: "中",
        date: "2026-07-21",
      },
      {
        name: "桂林文旅发展有限公司",
        event: "公司治理",
        source: "企业公告",
        summary: "完成董事会换届及管理层调整",
        sentiment: "中性",
        risk: "低",
        date: "2026-07-19",
      },
    ],
  },
  区域产业链客户画像: {
    unit: "家链上企业",
    empty: "配置区域、产业链与关系层级后生成客户画像",
    success: "产业链画像生成完成",
    push: "推送产业研究工作台",
    file: "区域产业链客户画像_演示.xls",
    fields: [
      "企业名称:name",
      "所在城市:city",
      "链属环节:chainRole",
      "核心产品:product",
      "上下游数量:relations",
      "景气评分:score",
      "融资需求:finance",
    ],
    rows: [
      {
        name: "长三角动力电池有限公司",
        city: "常州市",
        chainRole: "电芯制造",
        product: "动力电芯",
        relations: 28,
        score: 91,
        finance: "扩产融资",
      },
      {
        name: "甬创新材料有限公司",
        city: "宁波市",
        chainRole: "正极材料",
        product: "高镍正极",
        relations: 17,
        score: 86,
        finance: "流动资金",
      },
      {
        name: "嘉兴热管理科技有限公司",
        city: "嘉兴市",
        chainRole: "核心配套",
        product: "热管理系统",
        relations: 13,
        score: 82,
        finance: "技改融资",
      },
      {
        name: "无锡智能电驱有限公司",
        city: "无锡市",
        chainRole: "核心配套",
        product: "电驱系统",
        relations: 22,
        score: 89,
        finance: "研发融资",
      },
      {
        name: "湖州隔膜材料有限公司",
        city: "湖州市",
        chainRole: "关键材料",
        product: "湿法隔膜",
        relations: 16,
        score: 84,
        finance: "设备更新",
      },
      {
        name: "上海车规芯片有限公司",
        city: "上海市",
        chainRole: "上游核心",
        product: "车规级MCU",
        relations: 31,
        score: 93,
        finance: "研发及流动资金",
      },
      {
        name: "南京充换电科技有限公司",
        city: "南京市",
        chainRole: "基础设施",
        product: "充换电设备",
        relations: 19,
        score: 85,
        finance: "项目融资",
      },
      {
        name: "绍兴轻量化材料有限公司",
        city: "绍兴市",
        chainRole: "关键材料",
        product: "铝镁合金部件",
        relations: 14,
        score: 80,
        finance: "技改融资",
      },
      {
        name: "台州精密零部件有限公司",
        city: "台州市",
        chainRole: "核心配套",
        product: "轴承与齿轮",
        relations: 25,
        score: 83,
        finance: "订单融资",
      },
      {
        name: "苏州汽车软件有限公司",
        city: "苏州市",
        chainRole: "软件服务",
        product: "智能驾驶中间件",
        relations: 18,
        score: 88,
        finance: "研发融资",
      },
    ],
  },
  行政处罚与合规事项监测: {
    unit: "条合规事项",
    empty: "配置客户范围、事项类型与观察窗口后运行监测",
    success: "合规事项监测完成",
    push: "推送合规核验任务池",
    file: "行政处罚与合规监测结果_演示.xls",
    fields: [
      "企业名称:name",
      "事项类型:event",
      "处罚机关:authority",
      "处罚文号:caseNo",
      "处罚金额:amount",
      "整改状态:status",
      "合规等级:risk",
    ],
    rows: [
      {
        name: "海岳环保科技有限公司",
        event: "环保处罚",
        authority: "市生态环境局",
        caseNo: "环罚〔2026〕118号",
        amount: 320000,
        status: "整改中",
        risk: "中",
      },
      {
        name: "华南贸易有限公司",
        event: "税务事项",
        authority: "市税务局",
        caseNo: "税稽罚〔2026〕43号",
        amount: 180000,
        status: "已整改",
        risk: "低",
      },
      {
        name: "远航跨境物流有限公司",
        event: "海关处罚",
        authority: "海关",
        caseNo: "关缉罚〔2026〕76号",
        amount: 760000,
        status: "待复核",
        risk: "高",
      },
      {
        name: "华北医药流通有限公司",
        event: "市场监管处罚",
        authority: "省市场监督管理局",
        caseNo: "市监罚〔2026〕205号",
        amount: 450000,
        status: "整改中",
        risk: "中",
      },
      {
        name: "江城建设工程有限公司",
        event: "安全生产处罚",
        authority: "市应急管理局",
        caseNo: "应急罚〔2026〕91号",
        amount: 680000,
        status: "待复核",
        risk: "高",
      },
      {
        name: "青岛海产品贸易有限公司",
        event: "海关事项",
        authority: "青岛海关",
        caseNo: "关稽处〔2026〕38号",
        amount: 120000,
        status: "已整改",
        risk: "低",
      },
      {
        name: "西南数据科技有限公司",
        event: "数据合规",
        authority: "省网信主管部门",
        caseNo: "网信整改〔2026〕17号",
        amount: 0,
        status: "整改中",
        risk: "中",
      },
      {
        name: "杭州消费金融服务有限公司",
        event: "广告合规",
        authority: "市市场监督管理局",
        caseNo: "杭市监罚〔2026〕132号",
        amount: 210000,
        status: "已整改",
        risk: "低",
      },
      {
        name: "内蒙古能源技术有限公司",
        event: "环保处罚",
        authority: "自治区生态环境厅",
        caseNo: "蒙环罚〔2026〕64号",
        amount: 980000,
        status: "整改中",
        risk: "高",
      },
      {
        name: "大连国际物流有限公司",
        event: "外汇合规",
        authority: "外汇管理部门",
        caseNo: "汇检罚〔2026〕26号",
        amount: 350000,
        status: "待复核",
        risk: "中",
      },
    ],
  },
};

// 行内数据仅使用演示字段，正式环境由岗位权限与数据授权控制。
const internalDataOptions = [
  {
    id: "INT-CUSTOMER",
    name: "对公客户主数据",
    field: "客户关系",
    key: "customerRelation",
  },
  {
    id: "INT-OWNER",
    name: "客户经理管户关系",
    field: "管户机构",
    key: "ownerOrg",
  },
  {
    id: "INT-CREDIT",
    name: "授信与敞口信息",
    field: "授信敞口",
    key: "creditExposure",
  },
  {
    id: "INT-SETTLE",
    name: "账户与结算活跃度",
    field: "结算变化",
    key: "settlementTrend",
  },
  {
    id: "INT-PRODUCT",
    name: "存量产品持有情况",
    field: "产品覆盖",
    key: "productCoverage",
  },
  {
    id: "INT-TASK",
    name: "历史营销与风险任务",
    field: "历史反馈",
    key: "historyFeedback",
  },
];

const stableNumber = (text: string, mod = 100) =>
  Array.from(text).reduce((sum, ch) => sum + ch.charCodeAt(0), 0) % mod;
const expandRows = (rows: any[], target = 60) =>
  Array.from({ length: Math.max(target, rows.length) }, (_, i) => {
    const base = { ...rows[i % rows.length] };
    const round = Math.floor(i / rows.length);
    const no = i + 1;
    if (round === 0) return { ...base, key: no };
    const cityPool = [
      "北京市",
      "上海市",
      "深圳市",
      "成都市",
      "苏州市",
      "杭州市",
      "武汉市",
      "西安市",
      "青岛市",
      "合肥市",
      "宁波市",
      "厦门市",
    ];
    const suffix = [
      "智能科技",
      "先进材料",
      "精密装备",
      "数字产业",
      "绿色能源",
      "供应链服务",
    ][i % 6];
    return {
      ...base,
      key: no,
      name: `${cityPool[i % cityPool.length].replace("市", "")}${suffix}示范企业${String(no).padStart(2, "0")}号`,
      city: cityPool[i % cityPool.length],
      score:
        typeof base.score === "number"
          ? Math.max(35, Math.min(98, base.score - ((round * 3 + i) % 13) + 5))
          : base.score,
      amount:
        typeof base.amount === "number"
          ? Math.round(base.amount * (0.72 + (i % 9) * 0.07))
          : base.amount,
      date: `2026-${String(8 - Math.floor((i % 45) / 28)).padStart(2, "0")}-${String(1 + (i % 28)).padStart(2, "0")}`,
      caseNo: base.caseNo
        ? `${String(base.caseNo).replace(/号$/, "")}-${no}号`
        : base.caseNo,
    };
  });

const externalMetric = (resource: any, row: any, index: number) => {
  const seed = stableNumber(`${resource.id}-${row.name}-${index}`, 97);
  if (resource.name.includes("信用评分"))
    return `${650 + seed}分 / ${seed > 66 ? "低风险" : seed > 30 ? "中风险" : "关注"}`;
  if (resource.name.includes("招投标"))
    return `近一年${2 + (seed % 24)}次 / ${(300 + seed * 37).toLocaleString()}万元`;
  if (resource.name.includes("司法") || resource.name.includes("执行"))
    return seed > 72
      ? "存在高关注事项"
      : seed > 35
        ? "一般事项1项"
        : "未见重大风险";
  if (resource.name.includes("舆情"))
    return `${seed > 65 ? "负面" : seed > 28 ? "中性" : "正面"} / 热度${40 + seed}`;
  if (resource.name.includes("融资") || resource.name.includes("债券"))
    return `${seed % 3 ? "存在融资信号" : "暂无明确信号"} / ${(1000 + seed * 120).toLocaleString()}万元`;
  if (resource.name.includes("供应链"))
    return `集中度${25 + (seed % 61)}% / ${seed > 65 ? "偏高" : "正常"}`;
  if (resource.name.includes("ESG") || resource.name.includes("绿色"))
    return `${60 + (seed % 36)}分 / ${seed > 55 ? "绿色适配" : "一般"}`;
  if (resource.name.includes("用电") || resource.name.includes("活跃"))
    return `同比${seed % 2 ? "+" : "-"}${2 + (seed % 29)}%`;
  return `${resource.fields.split("、")[0] || "指标"}：${seed > 58 ? "较高" : seed > 25 ? "正常" : "偏低"}`;
};

const internalMetric = (key: string, row: any, index: number) => {
  const seed = stableNumber(`${key}-${row.name}-${index}`, 97);
  if (key === "customerRelation")
    return seed > 66
      ? "存量重点客户"
      : seed > 30
        ? "存量一般客户"
        : "潜在新客户";
  if (key === "ownerOrg")
    return [
      "北京分行公司部",
      "上海分行营业部",
      "四川分行公司部",
      "浙江分行营业部",
      "江苏分行公司部",
    ][seed % 5];
  if (key === "creditExposure")
    return seed > 25
      ? `${(800 + seed * 85).toLocaleString()}万元 / 使用率${35 + (seed % 61)}%`
      : "暂无授信";
  if (key === "settlementTrend")
    return `近90日${seed % 2 ? "增长" : "下降"}${3 + (seed % 28)}%`;
  if (key === "productCoverage")
    return `${1 + (seed % 5)}类 / ${seed > 54 ? "可交叉销售" : "基础覆盖"}`;
  return seed > 55
    ? "历史任务已跟进"
    : seed > 25
      ? "曾触达未转化"
      : "无历史任务";
};

function ProductInsightPanel({ product, rows, parameters = {} }: { product: string; rows: any[]; parameters?: Record<string, string> }) {
  const briefRef = useRef<HTMLDivElement>(null);
  const [briefOpen, setBriefOpen] = useState(false);
  const isIndustry = product.includes("产业") || product.includes("区域");
  const isRisk = product.includes("风险") || product.includes("司法") || product.includes("舆情") || product.includes("处罚");
  const topRows = rows.slice(0, 8);
  const parameterText = Object.entries(parameters).slice(0, 6).map(([k, v]) => `${k}：${v}`).join("　｜　");
  const regionLabel = parameters["地区"] || parameters["监测地域"] || parameters["项目地区"] || parameters["客户地域"] || "全国";
  const industryLabel = parameters["行业或产业链"] || parameters["行业"] || parameters["所属行业"] || "目标产业";
  const opportunityCount = rows.filter((x:any) => Number(x.opportunityIndex || x.score || 0) >= 75).length;
  const riskCount = rows.filter((x:any) => Number(x.riskIndex || 0) >= 60).length;
  const highConfidenceCount = rows.filter((x:any) => String(x.confidence || "").includes("高")).length;
  const averageOpportunity = rows.length ? Math.round(rows.reduce((sum:number, x:any) => sum + Number(x.opportunityIndex || x.score || 0), 0) / rows.length) : 0;
  const rankedOpportunity = [...rows].sort((a:any,b:any) => Number(b.opportunityIndex || b.score || 0) - Number(a.opportunityIndex || a.score || 0));
  const leadCompany = rankedOpportunity[0] || {};
  const secondCompany = rankedOpportunity[1] || leadCompany;
  const riskRows = rows.filter((x:any) => Number(x.riskIndex || 0) >= 60 || ["高", "中高"].includes(String(x.risk || "")));
  const leadRisk = [...riskRows].sort((a:any,b:any) => Number(b.riskIndex || 0) - Number(a.riskIndex || 0))[0];
  const opportunityAmount = rankedOpportunity.filter((x:any) => Number(x.opportunityIndex || x.score || 0) >= 75).reduce((sum:number,x:any) => sum + Number(x.amount || 0), 0);
  const factLabel = (row:any) => row?.event || row?.project || row?.qualification || "经营事件";
  const factDetail = (row:any) => row?.summary || row?.project || `${row?.city || regionLabel}${row?.industry || industryLabel}企业经营信号发生变化`;
  const amountText = (row:any) => Number(row?.amount || 0) > 0 ? `，涉及金额 ${Number(row.amount).toLocaleString()} 万元` : "";
  const insightTitle = isIndustry
    ? `${regionLabel}${industryLabel}景气度保持上行，机会集中于扩产、设备更新与核心配套环节`
    : isRisk
      ? `风险信号集中于少数关联主体，需对高敞口客户开展分级核验`
      : `新增线索与存量客户关系交叉后，形成可直接分派的三类营销优先级`;
  const insightPoints = isIndustry
    ? [
        `近一观察窗口识别 ${opportunityCount} 家高机会企业，平均机会指数 ${averageOpportunity}。`,
        `订单、招投标与扩产事件相互印证，核心部件和智能装备环节活跃度领先。`,
        `${riskCount} 家企业出现司法或经营风险抬升，建议在营销推进前完成风险复核。`,
      ]
    : isRisk
      ? [`共识别 ${riskCount} 家重点风险企业。`, "司法、舆情与经营异常构成主要风险来源。", "建议按敞口和关联传导范围分级处置并持续监测。"]
      : [`形成 ${opportunityCount} 家优先营销企业。`, "新增事件、客户关系和产品覆盖共同决定触达顺序。", "建议将空白客户优先推送客户经理并同步匹配融资方案。"];
  const facts:{title:string;fact:string;source:string;judgement:string}[] = isIndustry
    ? [
        {title:"直接业务信号",fact:`${leadCompany.name || "重点企业"}于 ${leadCompany.date || "本观察期"}出现“${factLabel(leadCompany)}”：${factDetail(leadCompany)}${amountText(leadCompany)}。`,source:"企业公告／招投标公告（演示）",judgement:`机会指数 ${leadCompany.opportunityIndex || leadCompany.score || "—"}，具备融资需求访谈依据`},
        {title:"同链交叉印证",fact:`${secondCompany.name || "同产业链企业"}同期出现“${factLabel(secondCompany)}”，与首要事件共同指向订单落地、产线建设或设备采购需求。`,source:"项目备案／企业经营事件（演示）",judgement:`${opportunityCount} 户达到机会阈值，相关事件金额合计 ${opportunityAmount.toLocaleString()} 万元`},
        leadRisk ? {title:"风险反证",fact:`${leadRisk.name}于 ${leadRisk.date || "本观察期"}出现“${factLabel(leadRisk)}”：${factDetail(leadRisk)}${amountText(leadRisk)}。`,source:"司法／处罚／舆情信息（演示）",judgement:`风险指数 ${leadRisk.riskIndex || "—"}，营销前应核验整改、敞口及关联传导`} : {title:"风险反证",fact:"本次命中企业未见达到重点风险阈值的司法、处罚或经营异常信号。",source:"司法／处罚／经营异常信息（演示）",judgement:"仍需在业务受理时执行最新风险查询"},
        {title:"可承接结论",fact:`本次按“${regionLabel}—${industryLabel}”口径分析 ${rows.length} 户，其中 ${opportunityCount} 户高机会、${riskCount} 户重点风险。`,source:"平台规则计算＋行内客户关系（演示）",judgement:"形成优先营销、人工核验和持续观察三类任务"},
      ]
    : isRisk
      ? [
          {title:"风险触发事实",fact:`${leadRisk?.name || leadCompany.name || "重点客户"}于 ${leadRisk?.date || leadCompany.date || "本观察期"}出现“${factLabel(leadRisk || leadCompany)}”${amountText(leadRisk || leadCompany)}。`,source:"司法裁判／执行／监管处罚信息（演示）",judgement:`风险指数 ${leadRisk?.riskIndex || leadCompany.riskIndex || "—"}，达到人工核验条件`},
          {title:"主体穿透结果",fact:`已按统一社会信用代码归并主体，并关联核验 ${rows.length} 户企业的股东、集团及关联企业。`,source:"工商登记＋股权关系＋集团客户信息（演示）",judgement:"避免同名误匹配，并识别风险向关联主体传导的可能性"},
          {title:"敞口核验条件",fact:`本次共有 ${riskCount} 户达到重点风险阈值，${highConfidenceCount} 项判断具备高置信度。`,source:"外部风险事件＋行内授信敞口（演示）",judgement:"先核验授信余额、担保关系和贷后检查情况，再决定处置等级"},
          {title:"分级处置结论",fact:"高风险客户进入当日核验，中风险客户进入持续监测，其余客户保留事件轨迹。",source:"风险规则 RISK-EXCLUDE-03（演示）",judgement:"系统生成任务，最终风险结论由风险人员确认"},
        ]
      : [
          {title:"线索触发",fact:`${leadCompany.name || "重点企业"}出现“${factLabel(leadCompany)}”，${factDetail(leadCompany)}${amountText(leadCompany)}。`,source:"资质／招投标／企业经营信息（演示）",judgement:`机会指数 ${leadCompany.opportunityIndex || leadCompany.score || "—"}，可进入客户经理核实环节`},
          {title:"主体与关系核验",fact:`已对 ${rows.length} 户企业补全工商主体、股权关系及行内客户关系。`,source:"工商登记＋行内客户主数据（演示）",judgement:"区分存量客户、一般客户和潜在新客户，避免重复触达"},
          {title:"机会分层依据",fact:`${opportunityCount} 户达到机会阈值，平均机会指数 ${averageOpportunity}，${riskCount} 户同时触发风险关注。`,source:"机会评分＋风险规则＋产品覆盖（演示）",judgement:"高机会低风险优先营销；机会与风险并存的先核验"},
          {title:"任务承接",fact:"名单已带出企业、触发事件、判断依据、建议产品和责任机构。",source:"平台任务编排结果（演示）",judgement:"由管户客户经理确认需求，回填触达结果和下一步安排"},
        ];
  const exportBrief = async () => {
    if (!briefRef.current) return;
    const close = message.loading("正在生成PPT版经营简报…", 0);
    try {
      const canvas = await html2canvas(briefRef.current, {
        backgroundColor: "#ffffff",
        scale: 2,
        useCORS: true,
        logging: false,
      });
      const link = document.createElement("a");
      link.download = `${product}_经营分析简报_演示.png`;
      link.href = canvas.toDataURL("image/png", 1);
      link.click();
      message.success("高清PNG简报已导出，可直接插入PPT");
    } catch {
      message.error("简报生成失败，请稍后重试");
    } finally {
      close();
    }
  };
  const downloadIndustryBrief = () => {
    const topCompanies = topRows.slice(0, 5).map((x:any, i:number) => `${i + 1}. ${x.name}：机会指数 ${x.opportunityIndex || x.score || "—"}，风险指数 ${x.riskIndex || "—"}，建议${x.action || "持续跟进"}`).join("<br/>");
    const html = `<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:KaiTi,serif;line-height:1.8;padding:42px;color:#222}h1{font-size:24px;text-align:center}h2{font-size:17px;border-left:4px solid #c7000b;padding-left:10px}p,li{font-size:14px}.meta{text-align:center;color:#666}.note{background:#f7f7f7;padding:12px}</style></head><body><h1>${regionLabel}${industryLabel}经营分析简报</h1><p class="meta">由“${product}”自动生成｜2026年8月｜演示数据</p><h2>一、核心结论</h2><p><b>${insightTitle}</b></p><ol>${insightPoints.map(x => `<li>${x}</li>`).join("")}</ol><h2>二、重点指标</h2><p>分析企业 ${rows.length} 家；高机会企业 ${opportunityCount} 家；重点风险企业 ${riskCount} 家；高置信度判断 ${highConfidenceCount} 项。</p><h2>三、重点企业与建议</h2><p>${topCompanies}</p><h2>四、业务建议</h2><p>建议将高机会、低风险企业优先推送客户经理，围绕扩产、设备更新、订单与供应链关系匹配融资方案；对风险抬升企业先行开展授信敞口和关联风险核验，并纳入持续监测。</p><p class="note">本材料由平台基于当前参数、内外部数据和通用规则自动形成，正式使用前应由业务人员复核。</p></body></html>`;
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([html], { type: "application/msword;charset=utf-8" }));
    link.download = `${regionLabel}${industryLabel}_行业简报_演示.doc`;
    link.click();
    URL.revokeObjectURL(link.href);
    message.success("行业简报已生成，可直接用于晨会、汇报或继续编辑");
  };
  const overview = <div className="insight-grid">
        <Card size="small" title={isIndustry ? "机会—风险分布" : isRisk ? "风险等级与敞口" : "客户价值—触达优先级"}>
          <ReactECharts style={{height:220}} option={{animation:false,tooltip:{},grid:{left:42,right:22,top:22,bottom:34},xAxis:{name:isRisk?"风险强度":"机会指数",min:20,max:100},yAxis:{name:isRisk?"授信敞口":"关系指数",min:10,max:100},series:[{type:"scatter",symbolSize:(v:any)=>12+v[2]/7,data:topRows.map((x:any,i:number)=>[Number(x.opportunityIndex)||45+i*6,Number(x.riskIndex)||28+i*7,Number(x.relationIndex)||50,x.name]),itemStyle:{color:(p:any)=>p.value[1]>70?"#c7000b":p.value[0]>75?"#1f8f55":"#d89614"},label:{show:true,formatter:(p:any)=>p.value[3],position:"top",fontSize:9}}]}} />
        </Card>
        <Card size="small" title={isIndustry ? "区域机会热度" : isRisk ? "风险来源构成" : "机会来源构成"}>
          <ReactECharts style={{height:220}} option={isIndustry?{animation:false,tooltip:{trigger:"axis"},grid:{left:76,right:24,top:24,bottom:30},xAxis:{type:"value"},yAxis:{type:"category",data:regionLabel === "全国" ? ["四川","江苏","浙江","广东","北京"] : [regionLabel,"周边区域A","周边区域B","同类区域均值"]},series:[{type:"bar",data:[92,86,81,77,68].slice(0,regionLabel === "全国" ? 5 : 4).map((v,i)=>Math.max(45,v-stableNumber(`${parameterText}-${i}`,10))),itemStyle:{color:"#c7000b",borderRadius:[0,4,4,0]}}]}:{animation:false,tooltip:{trigger:"item"},legend:{bottom:0},series:[{type:"pie",radius:[38,70],data:(isRisk?[{name:"司法",value:34},{name:"舆情",value:26},{name:"经营异常",value:22},{name:"供应链",value:18}]:[{name:"资质",value:28},{name:"招投标",value:31},{name:"扩产",value:24},{name:"产业协同",value:17}]),itemStyle:{borderColor:"#fff",borderWidth:2}}]}} />
        </Card>
      </div>;
  const relationTitle = isIndustry ? "产业链传导" : isRisk ? "风险关系图谱" : "企业关联网络";
  const relation = <ReactECharts style={{height:280}} option={{animation:false,tooltip:{},series:[{type:"graph",layout:"force",roam:true,label:{show:true,fontSize:10},force:{repulsion:260,edgeLength:[75,135]},data:[{name:product,symbolSize:72,itemStyle:{color:"#c7000b"}},...topRows.slice(0,6).map((x:any,i:number)=>({name:x.name,symbolSize:38+i%3*7,itemStyle:{color:i%3===0?"#d89614":"#3d6f91"}})),{name:isIndustry?"上游材料":"控股股东",symbolSize:42},{name:isIndustry?"核心部件":"关联企业",symbolSize:42}],links:topRows.slice(0,6).map((x:any)=>({source:product,target:x.name})).concat([{source:topRows[0]?.name,target:isIndustry?"上游材料":"控股股东"},{source:topRows[1]?.name,target:isIndustry?"核心部件":"关联企业"}])}]}} />;
  return <Card className="product-insight" title={<span><LineChartOutlined /> 产品业务成果中心</span>} extra={<Space><Tag color="green">持续更新</Tag><Button size="small" onClick={()=>setBriefOpen(true)}>预览行业简报</Button><Button size="small" icon={<DownloadOutlined />} onClick={exportBrief}>导出PPT简报</Button></Space>}>
    <div className="ppt-brief" ref={briefRef}>
      <div className="ppt-brief-head"><div><b>{product}</b><span>产品经营分析简报</span></div><small>生成时间：2026-08-09　｜　演示数据</small></div>
      {parameterText && <div className="ppt-parameter-strip"><b>本次运行参数</b><span>{parameterText}</span></div>}
      <div className="result-kpis">
        {[["分析企业", rows.length, "家"],["高机会企业", opportunityCount, "家"],["重点风险企业", riskCount, "家"],["平均机会指数", averageOpportunity, "分"],["高置信度判断", highConfidenceCount, "项"]].map((x:any)=><div key={x[0]}><small>{x[0]}</small><b>{x[1]}<em>{x[2]}</em></b></div>)}
      </div>
      <div className="ai-conclusion"><span className="ai-badge" aria-hidden="true">AI</span><div><small>AI综合研判 · 已引用本次参数与证据链</small><h2>{insightTitle}</h2><ul>{insightPoints.map(x=><li key={x}>{x}</li>)}</ul></div><Tag className="review-required-tag" color="red">人工复核后使用</Tag></div>
      <section><h3><span>01</span>经营总览</h3>{overview}</section>
      <div className="ppt-brief-middle">
        <section><h3><span>02</span>事实与证据链</h3><div className="evidence-chain evidence-chain-compact">{facts.map((x,i)=><button key={x.title}><span>0{i+1}</span><b>{x.title}</b><p>{x.fact}</p><small><strong>来源：</strong>{x.source}</small><small><strong>支持判断：</strong>{x.judgement}</small></button>)}</div></section>
        <section><h3><span>03</span>{relationTitle}</h3>{relation}</section>
      </div>
      <section><h3><span>04</span>研判与行动</h3><div className="decision-brief"><div><Tag color="red">有证据支持的业务判断</Tag><h3>{isIndustry?`${leadCompany.name || "重点企业"}等高机会主体具备融资需求核实价值，风险主体应从营销名单中单列`:isRisk?`${riskCount} 户达到重点风险阈值，处置前须完成主体、敞口和关联担保三项核验`:`${opportunityCount} 户达到营销阈值，应按客户关系与风险状态分层触达`}</h3><p>{isIndustry?`依据 ${leadCompany.name || "重点企业"}的“${factLabel(leadCompany)}”及 ${secondCompany.name || "同链企业"}的同期经营事件，可以合理判断相关企业可能产生项目建设、设备采购或订单周转融资需求；该判断仅支持发起需求访谈，不能直接替代授信结论。${leadRisk ? `${leadRisk.name}已出现“${factLabel(leadRisk)}”，应暂停自动营销并先核验整改进度、行内敞口及关联影响。` : "本期未命中重点风险阈值，但受理前仍需刷新风险信息。"}`:isRisk?`${leadRisk?.name || "首要风险客户"}的“${factLabel(leadRisk || leadCompany)}”构成直接风险触发；外部事件只能说明风险线索，风险人员需结合授信余额、担保链、回款和贷后检查记录确定是否压降、观察或维持。`:`${leadCompany.name || "首要线索企业"}的“${factLabel(leadCompany)}”构成触达理由；对存量客户由管户经理核实融资需求，对空白客户分派属地机构，对风险指数偏高的企业先完成合规与风险核验。`}</p><div className="judgement-boundary"><SafetyCertificateOutlined /><span><b>判断边界</b> 平台给出线索优先级和核验方向；客户需求、授信可行性与风险处置结论均须由对应岗位确认。</span></div></div><div className="action-cards">{["发起需求访谈","生成授信核验","加入持续监测","导出证据明细"].map((x,i)=><button key={x} onClick={()=>message.success(`${x}任务已生成`)}><span>0{i+1}</span><b>{x}</b><small>{[`${Math.max(0,opportunityCount-riskCount)}户高机会低风险`,`${riskCount}户需先行复核`,`${Math.max(0,rows.length-opportunityCount)}户保留观察`,`${rows.length}户含来源与日期`][i]}</small></button>)}</div></div></section>
      <section className="office-output"><h3><span>05</span>办公材料中心</h3><p>分析结果自动转化为不同岗位可直接使用的工作材料，减少重复取数、整理、撰写和任务分派。</p><div>{[["行业分析简报","Word","含观点、指标、重点企业与建议"],["晨会要点","1页","供经营分析与部门晨会使用"],["客户营销清单","Excel","含负责人、触达话术与产品建议"],["风险核验任务","任务单","推送授信与风险岗位闭环处理"]].map((x:any,i:number)=><button key={x[0]} onClick={i===0?downloadIndustryBrief:()=>message.success(`${x[0]}已生成并进入待办中心`)}><FileDoneOutlined /><span><b>{x[0]}</b><small>{x[2]}</small></span><Tag>{x[1]}</Tag></button>)}</div></section>
      <div className="ppt-brief-foot">资料来源：中国工商银行外部数据产品管理平台　｜　本简报仅用于内部演示</div>
    </div>
    <Modal width={840} open={briefOpen} onCancel={()=>setBriefOpen(false)} title={`${regionLabel}${industryLabel}经营分析简报`} footer={<Space><Button onClick={()=>setBriefOpen(false)}>关闭</Button><Button type="primary" icon={<DownloadOutlined />} onClick={downloadIndustryBrief}>下载Word简报</Button></Space>}>
      <div className="brief-preview"><div className="brief-cover"><Tag color="red">自动生成</Tag><h2>{insightTitle}</h2><p>基于当前运行参数、外数资源、行内客户关系与通用规则综合形成</p></div><h3>核心观点</h3><ol>{insightPoints.map(x=><li key={x}>{x}</li>)}</ol><h3>重点企业</h3><Table size="small" pagination={false} rowKey={(x:any)=>x.name} dataSource={topRows.slice(0,5)} columns={[{title:"企业",dataIndex:"name"},{title:"机会指数",render:(_:any,x:any)=>x.opportunityIndex||x.score||"—"},{title:"风险指数",dataIndex:"riskIndex"},{title:"建议动作",dataIndex:"action"}]} /><div className="brief-audit"><SafetyCertificateOutlined /> 内容已标注数据口径与生成依据，须经业务人员复核后正式使用。</div></div>
    </Modal>
  </Card>;
}

function Workbench({ go }: { go: (v: View) => void }) {
  const publishedRecord =
    typeof window !== "undefined"
      ? JSON.parse(localStorage.getItem("icbc-published-product") || "null")
      : null;
  const canonicalProduct = "全国产业链机会与风险识别";
  const productCatalog =
    publishedRecord?.name && !workbenchProducts[publishedRecord.name]
      ? {
          ...workbenchProducts,
          [publishedRecord.name]: workbenchProducts[canonicalProduct],
        }
      : workbenchProducts;
  const runtimeCatalog =
    publishedRecord?.name && !workbenchRuntime[publishedRecord.name]
      ? {
          ...workbenchRuntime,
          [publishedRecord.name]: workbenchRuntime[canonicalProduct],
        }
      : workbenchRuntime;
  const [generated, setGenerated] = useState(false);
  const [loading, setLoading] = useState(false);
  const [pushed, setPushed] = useState(false);
  const [workMode, setWorkMode] = useState<"static" | "product">("static");
  const [objectInput, setObjectInput] = useState("");
  const [selectedObjects, setSelectedObjects] = useState<string[]>([
    "华北智造装备有限公司",
    "京芯半导体设备股份有限公司",
    "北方精密传动科技有限公司",
    "中关村工业软件有限公司",
  ]);
  const [objectSource, setObjectSource] = useState("手工添加");
  const objectLists: Record<string, string[]> = {
    专精特新企业营销名单: [
      "中关村工业软件有限公司",
      "京芯半导体设备股份有限公司",
      "北方精密传动科技有限公司",
      "华北智造装备有限公司",
    ],
    先进制造产业链机会名单: [
      "成渝数控机床有限公司",
      "蜀航精密制造有限公司",
      "蓉芯功率半导体有限公司",
      "川西工业机器人有限公司",
    ],
    重大项目中标客户名单: [
      "华东智能装备集团有限公司",
      "苏州精工自动化有限公司",
      "南京新材料科技股份有限公司",
    ],
    存量客户风险关注名单: [
      "远航供应链有限公司",
      "恒通精密部件有限公司",
      "东部新能材料有限公司",
    ],
  };
  const [product, setProduct] = useState(Object.keys(productCatalog)[0]);
  const [selectedResources, setSelectedResources] = useState<string[]>(
    productCatalog[Object.keys(productCatalog)[0]].resources,
  );
  const [selectedInternal, setSelectedInternal] = useState<string[]>([
    "INT-CUSTOMER",
    "INT-OWNER",
    "INT-CREDIT",
    "INT-SETTLE",
  ]);
  const [selectedAnalysisProducts, setSelectedAnalysisProducts] = useState<string[]>([
    "企业司法风险监测",
    "区域产业链客户画像",
  ]);
  const [resourceOpen, setResourceOpen] = useState(false);
  const [resourceQuery, setResourceQuery] = useState("");
  const [conditionValues, setConditionValues] = useState<
    Record<string, string>
  >({});
  const fieldCatalog = (name: string) =>
    (runtimeCatalog[name]?.fields || []).map((x: string) => {
      const [label, key] = x.split(":");
      return { key, label, help: `${name}正式输出字段：${label}` };
    });
  const [selectedFields, setSelectedFields] = useState<string[]>([
    ...fieldCatalog(Object.keys(productCatalog)[0]).map((x) => x.key),
    "customerRelation",
    "creditExposure",
    "settlementTrend",
    "previousOpportunity",
    "opportunityIndex",
    "opportunityChange",
    "riskIndex",
    "relationIndex",
    "confidence",
    "decision",
    "decisionBasis",
    "action",
  ]);
  const [catalogTab, setCatalogTab] = useState("resources");
  const config = productCatalog[product];
  const runtime = runtimeCatalog[product];
  const isEventProduct = [
    "企业司法风险监测",
    "舆情风险预警",
    "行政处罚与合规事项监测",
  ].includes(product);
  const baseResourceIds = productCatalog[product].resources;
  const addedResources = resources.filter(
    (r: any) =>
      selectedResources.includes(r.id) && !baseResourceIds.includes(r.id),
  );
  const externalFields = addedResources.map((r: any) => ({
    key: `ext_${r.id.replaceAll("-", "_")}`,
    label: `${r.name}指标`,
    help: `新增外数资源 ${r.id} 进入结果与评分模型`,
  }));
  const internalFields = internalDataOptions
    .filter((x) => selectedInternal.includes(x.id))
    .map((x) => ({
      key: x.key,
      label: x.field,
      help: `行内数据：${x.name}（按岗位授权使用）`,
    }));
  const decisionFields = [
    {
      key: "previousOpportunity",
      label: "上期机会指数",
      help: "最近一次有效评估的机会状态",
    },
    {
      key: "opportunityIndex",
      label: "本期机会指数",
      help: "产业事件与内外数特征共同更新",
    },
    {
      key: "opportunityChange",
      label: "机会变化",
      help: "本期相对比较基期的变化",
    },
    {
      key: "riskIndex",
      label: "风险指数",
      help: "风险事件、基本面与敞口的综合结果",
    },
    {
      key: "relationIndex",
      label: "客户关系指数",
      help: "存贷款、结算、产品与管户关系的综合刻画",
    },
    {
      key: "confidence",
      label: "判断置信度",
      help: "数据完整度、一致性与模型稳定性",
    },
    {
      key: "decision",
      label: "综合判断",
      help: "三维指数与规则组合形成的业务判断",
    },
    {
      key: "decisionBasis",
      label: "变化原因与判断依据",
      help: "展示新增事件和关键内外数特征的贡献",
    },
    {
      key: "action",
      label: "建议动作",
      help: "营销、授信、监测或人工复核动作",
    },
  ];
  const fields = [
    ...fieldCatalog(product),
    ...externalFields,
    ...internalFields,
    ...decisionFields,
  ].filter(
    (x: any, i: number, a: any[]) => a.findIndex((y) => y.key === x.key) === i,
  );
  const resultRows = expandRows(runtime.rows, 180).map(
    (row: any, index: number) => {
      const baseScore =
        typeof row.score === "number"
          ? row.score
          : row.risk === "高"
            ? 42
            : row.risk === "中"
              ? 65
              : 82;
      const extValues = Object.fromEntries(
        addedResources.map((r: any) => [
          `ext_${r.id.replaceAll("-", "_")}`,
          externalMetric(r, row, index),
        ]),
      );
      const intValues = Object.fromEntries(
        internalDataOptions
          .filter((x) => selectedInternal.includes(x.id))
          .map((x) => [x.key, internalMetric(x.key, row, index)]),
      );
      const extAdjustment = addedResources.reduce(
        (sum: any, r: any) =>
          sum + (stableNumber(`${r.id}-${row.name}`, 11) - 5),
        0,
      );
      const internalAdjustment = selectedInternal.reduce(
        (sum, id) => sum + (stableNumber(`${id}-${row.name}`, 7) - 3),
        0,
      );
      const riskPenalty = row.risk === "高" ? -10 : row.risk === "中" ? -4 : 0;
      const scoreAdjustment = Math.max(
        -18,
        Math.min(18, extAdjustment + internalAdjustment + riskPenalty),
      );
      const fusionScore = Math.max(
        0,
        Math.min(100, baseScore + scoreAdjustment),
      );
      const opportunityIndex = fusionScore;
      const previousOpportunity = Math.max(
        0,
        Math.min(
          100,
          opportunityIndex - (stableNumber(`${row.name}-change`, 23) - 9),
        ),
      );
      const opportunityChange = opportunityIndex - previousOpportunity;
      const riskIndex = Math.max(
        8,
        Math.min(
          96,
          (row.risk === "高" ? 78 : row.risk === "中" ? 56 : 28) +
            stableNumber(`${row.name}-risk`, 17) -
            8,
        ),
      );
      const relationIndex = selectedInternal.length
        ? Math.max(
            18,
            Math.min(
              96,
              36 +
                selectedInternal.length * 7 +
                stableNumber(`${row.name}-relation`, 29),
            ),
          )
        : 0;
      const confidence =
        selectedResources.length >= 7 && selectedInternal.length >= 4
          ? "高"
          : selectedResources.length >= 5
            ? "中高"
            : "中";
      const decision =
        riskIndex >= 75
          ? "风险升级处置"
          : opportunityIndex >= 82 && relationIndex < 55
            ? "重点拓展"
            : opportunityIndex >= 82
              ? "优先营销"
              : opportunityIndex >= 68 && riskIndex < 60
                ? "持续跟进"
                : "人工复核";
      const basis = [
        row.event &&
          `${row.event}触发${opportunityChange >= 0 ? "机会上调" : "审慎调整"}${Math.abs(opportunityChange)}分`,
        addedResources[0]?.name && `${addedResources[0].name}参与校准`,
        selectedInternal.includes("INT-CREDIT") && "授信敞口已核验",
        selectedInternal.includes("INT-SETTLE") && "结算趋势已纳入",
        `风险指数${riskIndex}`,
      ]
        .filter(Boolean)
        .join("；");
      const action =
        decision === "风险升级处置"
          ? "推送风险人员核验敞口与处置方案"
          : decision === "重点拓展"
            ? "生成潜客拓展任务并匹配融资方案"
            : decision === "优先营销"
              ? "推送管户客户经理开展融资需求访谈"
              : decision === "持续跟进"
                ? "纳入观察名单并在下一事件触发时重评"
                : "生成跨岗位人工复核任务";
      return {
        ...row,
        ...extValues,
        ...intValues,
        baseScore,
        scoreAdjustment: `${scoreAdjustment >= 0 ? "+" : ""}${scoreAdjustment}`,
        fusionScore,
        score: fusionScore,
        previousOpportunity,
        opportunityIndex,
        opportunityChange: `${opportunityChange >= 0 ? "+" : ""}${opportunityChange}`,
        riskIndex,
        relationIndex: relationIndex || "未导入内数",
        confidence,
        decision,
        decisionBasis: basis,
        action,
      };
    },
  );
  const activeParameters = Object.fromEntries(
    config.conditions.map((x: any) => [x.label, conditionValues[x.key] || x.values[0]]),
  );
  const requestedLimit = Number(String(conditionValues.limit || "").match(/\d+/)?.[0] || 180);
  const opportunityFloor = Number(String(conditionValues.opportunity || "").match(/\d+/)?.[0] || 0);
  const riskCeiling = String(conditionValues.risk || "").includes("40") ? 40 : String(conditionValues.risk || "").includes("60") ? 60 : 100;
  const effectiveRows = resultRows
    .map((row: any, index: number) => ({
      ...row,
      city:
        conditionValues.region && conditionValues.region !== "全国"
          ? conditionValues.region
          : row.city,
      industry: conditionValues.industry || row.industry,
      chainRole:
        conditionValues.chain && conditionValues.chain !== "全产业链"
          ? conditionValues.chain
          : row.chainRole,
      parameterMatch: `${conditionValues.region || "全国"} · ${conditionValues.industry || "全部行业"} · ${conditionValues.window || "最新窗口"}`,
      opportunityIndex: Math.max(0, Math.min(100, Number(row.opportunityIndex) + stableNumber(`${JSON.stringify(conditionValues)}-${row.name}`, 9) - 4)),
      key: row.key || index + 1,
    }))
    .filter((row: any) => Number(row.opportunityIndex || row.score || 0) >= opportunityFloor && Number(row.riskIndex || 0) <= riskCeiling)
    .slice(0, requestedLimit);
  useEffect(() => {
    const raw = localStorage.getItem("icbc-active-plan");
    if (!raw) return;
    try {
      const plan = JSON.parse(raw);
      if (productCatalog[plan.product]) {
        setProduct(plan.product);
        setSelectedResources(
          plan.resources || productCatalog[plan.product].resources,
        );
        setSelectedFields(
          plan.fields || fieldCatalog(plan.product).map((x) => x.key),
        );
        const byKey: Record<string, string> = {};
        productCatalog[plan.product].conditions.forEach((x: any) => {
          if (plan.conditions?.[x.label])
            byKey[x.key] = plan.conditions[x.label];
        });
        setConditionValues(byKey);
        message.success(`已载入个人方案：${plan.name}`);
      }
    } finally {
      localStorage.removeItem("icbc-active-plan");
    }
  }, []);
  useEffect(() => {
    const selected = localStorage.getItem("icbc-workbench-product");
    if (selected && productCatalog[selected]) {
      changeProduct(selected);
      localStorage.removeItem("icbc-workbench-product");
    }
  }, []);
  const changeProduct = (value: string) => {
    setProduct(value);
    setSelectedResources([...productCatalog[value].resources]);
    setSelectedInternal([
      "INT-CUSTOMER",
      "INT-OWNER",
      "INT-CREDIT",
      "INT-SETTLE",
    ]);
    setSelectedFields([
      ...fieldCatalog(value).map((x) => x.key),
      "customerRelation",
      "creditExposure",
      "settlementTrend",
      ...(value === canonicalProduct
        ? [
            "previousOpportunity",
            "opportunityIndex",
            "opportunityChange",
            "riskIndex",
            "relationIndex",
            "confidence",
            "decision",
            "decisionBasis",
            "action",
          ]
        : ["fusionScore", "decision", "decisionBasis"]),
    ]);
    setConditionValues({});
    setGenerated(false);
    setPushed(false);
  };
  const savePlan = () => {
    const plans = JSON.parse(
      localStorage.getItem("icbc-personal-plans") || "[]",
    );
    const plan = {
      id: `PLAN-${Date.now()}`,
      name: `${product}·个人方案`,
      product,
      resources: selectedResources,
      fields: selectedFields,
      conditions: Object.fromEntries(
        config.conditions.map((x: any) => [
          x.label,
          conditionValues[x.key] || x.values[0],
        ]),
      ),
      updated: new Date().toLocaleString("zh-CN", { hour12: false }),
    };
    localStorage.setItem(
      "icbc-personal-plans",
      JSON.stringify([plan, ...plans].slice(0, 10)),
    );
    message.success("个人方案已保存，可在“我的关注—个人方案”中调取");
  };
  const run = () => {
    setLoading(true);
    setTimeout(() => {
      setLoading(false);
      setGenerated(true);
      message.success(
        `${runtime.success}，当前参数命中 ${effectiveRows.length} ${runtime.unit}`,
      );
    }, 900);
  };
  const addObject = () => {
    const name = objectInput.trim();
    if (!name) return;
    setSelectedObjects((v) => [...new Set([...v, name])]);
    setObjectInput("");
    setGenerated(false);
  };
  const applyObjectList = () => {
    const names = objectLists[objectSource] || [];
    setSelectedObjects((v) => [...new Set([...v, ...names])]);
    setGenerated(false);
    message.success(`已从“${objectSource}”加入 ${names.length} 个研究对象`);
  };
  const shownResources = resources.filter(
    (r: any) =>
      `${r.id}${r.name}${r.supplier}${r.fields}`
        .toLowerCase()
        .includes(resourceQuery.toLowerCase()) && r.auth !== "采购中",
  );
  const resultColumns = fields
    .filter((x) => selectedFields.includes(x.key))
    .map((x: any, i: number) => ({
      title: x.label,
      dataIndex: x.key,
      fixed: i === 0 ? ("left" as const) : undefined,
      width:
        i === 0
          ? 220
          : ["decisionBasis", "creditExposure", "action"].includes(x.key)
            ? 260
            : undefined,
      render: (v: any) =>
        x.key === "risk" ? (
          <Tag color={v === "高" ? "red" : v === "中" ? "orange" : "green"}>
            {v}
          </Tag>
        ) : x.key === "decision" ? (
          <Tag
            color={
              String(v).includes("优先") || String(v).includes("拓展")
                ? "green"
                : String(v).includes("风险")
                  ? "red"
                  : "orange"
            }
          >
            {v}
          </Tag>
        ) : x.key === "opportunityChange" ? (
          <b className={Number(v) >= 0 ? "positive" : "red"}>{v}</b>
        ) : x.key === "amount" && typeof v === "number" ? (
          `¥${v.toLocaleString()} 万元`
        ) : [
            "score",
            "fusionScore",
            "opportunityIndex",
            "riskIndex",
            "relationIndex",
          ].includes(x.key) ? (
          <b className="red">{v}</b>
        ) : (
          (v ?? "—")
        ),
    }));
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            产品工作台 <DemoTag />
          </h1>
          <p>
            先按企业名称或统一社会信用代码定位主体，再选取外数资源与标准产品，生成企业画像、关系网络、风险事件和综合研判
          </p>
        </div>
        <Space>
          <Button icon={<HeartOutlined />} onClick={savePlan}>
            保存研究方案
          </Button>
          <Button onClick={() => message.success("新增产品建议已提交")}>
            提交新增产品建议
          </Button>
        </Space>
      </div>
      <Card className="wb-mode-switch">
        <div>
          <b>企业综合查询</b>
          <span>以企业为起点组合资源与产品，复杂产业任务可切换至标准产品分析</span>
        </div>
        <Radio.Group
          optionType="button"
          buttonStyle="solid"
          value={workMode}
          onChange={(e) => {
            setWorkMode(e.target.value);
            setGenerated(false);
          }}
          options={[
            { label: "企业综合查询", value: "static" },
            { label: "标准产品分析", value: "product" },
          ]}
        />
        <div className="wb-mode-value">
          <b>
            {workMode === "static"
              ? "适合单户或批量企业查询、尽调与画像"
              : "适合产业事件研判与持续经营"}
          </b>
          <span>
            {workMode === "static"
              ? "名称/代码定位主体，按需调用外数和产品"
              : "事件触发、模型计算、任务推送、效果反馈"}
          </span>
        </div>
      </Card>
      <div className="workbench-note">
        <SafetyCertificateOutlined />
        <span>
          <b>{workMode === "static" ? "静态研究空间" : "业务侧受控组合"}</b>　
          {workMode === "static"
            ? "仅调用已授权字段；名单可由手工添加、文件导入或正式产品结果生成，查询本身不改变正式产品口径。"
            : "可选择开放资源、输出字段和业务参数；主体关联、指标算法、风险口径及审批版本由正式产品统一控制。"}
        </span>
      </div>
      {workMode === "static" && (
        <Card
          className="wb-object-pool"
          title="1. 建立研究对象池"
          extra={
            <Space>
              <Tag color="red">{selectedObjects.length} 个对象</Tag>
              <Button
                size="small"
                danger
                disabled={!selectedObjects.length}
                onClick={() => {
                  setSelectedObjects([]);
                  setGenerated(false);
                }}
              >
                清空
              </Button>
            </Space>
          }
        >
          <div className="object-source-grid">
            <div>
              <b>直接添加企业或主体</b>
              <Space.Compact block>
                <AutoComplete
                  value={objectInput}
                  onChange={setObjectInput}
                  onSelect={(v: string) => {
                    setSelectedObjects((x) => [...new Set([...x, v])]);
                    setObjectInput("");
                  }}
                  options={[
                    "华北智造装备有限公司",
                    "京芯半导体设备股份有限公司",
                    "成渝数控机床有限公司",
                    "华东智能装备集团有限公司",
                  ].map((value) => ({ value }))}
                  placeholder="输入企业名称或统一社会信用代码"
                />
                <Button
                  type="primary"
                  icon={<PlusOutlined />}
                  onClick={addObject}
                >
                  添加
                </Button>
              </Space.Compact>
            </div>
            <div>
              <b>引用产品加工出的名单</b>
              <Space.Compact block>
                <Select
                  value={objectSource}
                  onChange={setObjectSource}
                  options={Object.keys(objectLists).map((value) => ({ value }))}
                />
                <Button onClick={applyObjectList}>加入对象池</Button>
              </Space.Compact>
            </div>
            <div>
              <b>批量导入研究对象</b>
              <Button
                block
                icon={<CloudUploadOutlined />}
                onClick={() => {
                  setSelectedObjects((v) => [
                    ...new Set([
                      ...v,
                      "沪科工业视觉有限公司",
                      "长三角精密仪器有限公司",
                      "江北智能传感有限公司",
                    ]),
                  ]);
                  message.success("已模拟导入3个对象，并完成主体去重与标准化");
                }}
              >
                导入 Excel／粘贴名单
              </Button>
            </div>
          </div>
          <div className="object-tags">
            {selectedObjects.length ? (
              selectedObjects.map((x) => (
                <Tag
                  key={x}
                  closable
                  onClose={() => {
                    setSelectedObjects((v) => v.filter((n) => n !== x));
                    setGenerated(false);
                  }}
                >
                  {x}
                </Tag>
              ))
            ) : (
              <Empty
                image={Empty.PRESENTED_IMAGE_SIMPLE}
                description="请添加研究对象，或引用已有产品名单"
              />
            )}
          </div>
          <div className="object-pool-foot">
            <span>
              <CheckCircleFilled /> 自动完成主体匹配、名称标准化和重复对象合并
            </span>
            <Button
              size="small"
              onClick={() => message.success("对象池已保存至我的关注")}
            >
              保存为自选对象组
            </Button>
          </div>
        </Card>
      )}
      {workMode === "product" && (
        <Card className="wb-product-guide">
          <span className="wb-step-number">1</span>
          <div>
            <small>选择正式产品</small>
            <Select
              value={product}
              onChange={changeProduct}
              options={Object.keys(productCatalog).map((value) => ({ value }))}
            />
          </div>
          <div className="wb-product-boundary">
            <Tag color="red">
              正式产品{" "}
              {product === publishedRecord?.name
                ? publishedRecord?.version || "V3.0"
                : "V1.0"}
            </Tag>
            <b>{config.abilities.length}项开放能力</b>
            <span>核心加工链路已核准，业务人员仅配置开放项</span>
          </div>
        </Card>
      )}
      <div className="wind-workbench">
        <Card className="wind-catalog" title="2. 可选内容">
          <Tabs
            activeKey={catalogTab}
            onChange={setCatalogTab}
            items={[
              {
                key: "resources",
                label: `外数 ${selectedResources.length}`,
                children: (
                  <>
                    <Input
                      size="small"
                      prefix={<SearchOutlined />}
                      placeholder="搜索已授权外数资源"
                      value={resourceQuery}
                      onChange={(e) => setResourceQuery(e.target.value)}
                    />
                    <div className="wind-option-list">
                      {shownResources.slice(0, 12).map((r: any) => (
                        <label key={r.id}>
                          <Checkbox
                            checked={selectedResources.includes(r.id)}
                            onChange={(e) => {
                              setSelectedResources((v) =>
                                e.target.checked
                                  ? [...new Set([...v, r.id])]
                                  : v.filter((x) => x !== r.id),
                              );
                              if (
                                e.target.checked &&
                                !baseResourceIds.includes(r.id)
                              )
                                setSelectedFields((v) => [
                                  ...new Set([
                                    ...v,
                                    `ext_${r.id.replaceAll("-", "_")}`,
                                    "fusionScore",
                                    "decision",
                                    "decisionBasis",
                                  ]),
                                ]);
                              setGenerated(false);
                            }}
                          />
                          <span>
                            <b>{r.name}</b>
                            <small>
                              {r.type} · {r.frequency}
                            </small>
                          </span>
                        </label>
                      ))}
                    </div>
                    <Button
                      block
                      size="small"
                      icon={<PlusOutlined />}
                      onClick={() => setResourceOpen(true)}
                    >
                      浏览完整外数资源库
                    </Button>
                  </>
                ),
              },
              {
                key: "products",
                label: `产品 ${selectedAnalysisProducts.length}`,
                children: (
                  <>
                    <div className="internal-data-note">产品提供已核准的指标口径、加工逻辑与判断结果，可与单项外数共同形成企业综合分析。</div>
                    <div className="wind-option-list">
                      {initialProducts.slice(0, 8).map((x: any) => (
                        <label key={x.id}>
                          <Checkbox checked={selectedAnalysisProducts.includes(x.name)} onChange={(e) => {
                            setSelectedAnalysisProducts((v) => e.target.checked ? [...new Set([...v, x.name])] : v.filter((n) => n !== x.name));
                            setGenerated(false);
                          }} />
                          <span><b>{x.name}</b><small>{x.category} · {x.scope} · 输出可解释结论</small></span>
                        </label>
                      ))}
                    </div>
                  </>
                ),
              },
              {
                key: "internal",
                label: `内数 ${selectedInternal.length}`,
                children: (
                  <>
                    <div className="internal-data-note">
                      行内数据仅在授权范围内参与计算，演示环境不展示真实客户数据。
                    </div>
                    <div className="wind-option-list">
                      {internalDataOptions.map((x) => (
                        <label key={x.id}>
                          <Checkbox
                            checked={selectedInternal.includes(x.id)}
                            onChange={(e) => {
                              setSelectedInternal((v) =>
                                e.target.checked
                                  ? [...v, x.id]
                                  : v.filter((id) => id !== x.id),
                              );
                              if (e.target.checked)
                                setSelectedFields((v) => [
                                  ...new Set([
                                    ...v,
                                    x.key,
                                    "fusionScore",
                                    "decision",
                                    "decisionBasis",
                                  ]),
                                ]);
                              setGenerated(false);
                            }}
                          />
                          <span>
                            <b>{x.name}</b>
                            <small>行内受控 · 输出字段：{x.field}</small>
                          </span>
                        </label>
                      ))}
                    </div>
                  </>
                ),
              },
              {
                key: "fields",
                label: `字段 ${selectedFields.length}`,
                children: (
                  <div className="wind-option-list">
                    {fields.map((x) => (
                      <label key={x.key}>
                        <Checkbox
                          checked={selectedFields.includes(x.key)}
                          onChange={(e) =>
                            setSelectedFields((v) =>
                              e.target.checked
                                ? [...v, x.key]
                                : v.filter((k) => k !== x.key),
                            )
                          }
                        />
                        <span>
                          <b>{x.label}</b>
                          <small>{x.help}</small>
                        </span>
                      </label>
                    ))}
                  </div>
                ),
              },
              {
                key: "conditions",
                label: `条件 ${config.conditions.length}`,
                children: (
                  <div className="wind-condition-index">
                    {config.conditions.map((x: any, i: number) => (
                      <button
                        key={x.key}
                        onClick={() => {
                          const el = document.getElementById(
                            `condition-${x.key}`,
                          );
                          el?.scrollIntoView({
                            behavior: "smooth",
                            block: "center",
                          });
                        }}
                      >
                        <span>{i + 1}</span>
                        <b>{x.label}</b>
                        <small>{conditionValues[x.key] || x.values[0]}</small>
                      </button>
                    ))}
                  </div>
                ),
              },
            ]}
          />
        </Card>
        <Card
          className="wind-config"
          title={workMode === "static" ? "3. 设置静态筛选条件" : "3. 参数配置"}
          extra={<Tag>{config.conditions.length}项可调</Tag>}
        >
          <Form layout="vertical">
            <div className="wind-config-grid">
              {config.conditions.map((x: any) => (
                <Form.Item
                  id={`condition-${x.key}`}
                  label={x.label}
                  key={x.key}
                >
                  <Select
                    value={conditionValues[x.key] || x.values[0]}
                    onChange={(value) => {
                      setConditionValues((v) => ({ ...v, [x.key]: value }));
                      setGenerated(false);
                    }}
                    options={x.values.map((value: string) => ({ value }))}
                  />
                </Form.Item>
              ))}
            </div>
            {product.includes("专精特新") && (
              <Checkbox defaultChecked>优先展示国资控股企业</Checkbox>
            )}
          </Form>
          <div className="ability-boundary">
            <b>
              {workMode === "static" ? "查询时自动调用" : "后台固定加工能力"}
            </b>
            {(workMode === "static"
              ? ["主体标准化", "字段权限校验", "条件筛选", "结果去重"]
              : config.abilities
            ).map((x: string) => (
              <Tag key={x} icon={<CheckCircleFilled />}>
                {x}
              </Tag>
            ))}
          </div>
        </Card>
        <Card className="wind-summary" title="4. 当前组合">
          <div className="summary-product">
            <small>
              {workMode === "static" ? "研究对象池" : "基于正式产品"}
            </small>
            <b>
              {workMode === "static"
                ? `${selectedObjects.length} 个对象 · ${objectSource}`
                : product}
            </b>
          </div>
          <div className="fusion-summary">
            <span>
              <b>{selectedResources.length}</b>项外数
            </span>
            <em>＋</em>
            <span>
              <b>{selectedInternal.length}</b>项内数
            </span>
            <em>→</em>
            <span>
              <b>
                {workMode === "static"
                  ? "静态筛选表"
                  : product === canonicalProduct
                    ? "三维动态评估"
                    : "融合判断"}
              </b>
            </span>
          </div>
          <div className="summary-block">
            <span>
              外数资源与指标 <Tag>{selectedResources.length}</Tag>
            </span>
            {resources
              .filter((r: any) => selectedResources.includes(r.id))
              .slice(0, 3)
              .map((x: any) => (
                <p key={x.id}>
                  <DatabaseOutlined />
                  {x.name}
                </p>
              ))}
          </div>
          <div className="summary-block">
            <span>调用产品 <Tag>{selectedAnalysisProducts.length}</Tag></span>
            {selectedAnalysisProducts.slice(0, 3).map((x) => (
              <p key={x}><ProductOutlined />{x}</p>
            ))}
          </div>
          <div className="summary-block">
            <span>
              行内数据 <Tag>{selectedInternal.length}</Tag>
            </span>
            {internalDataOptions
              .filter((x) => selectedInternal.includes(x.id))
              .slice(0, 3)
              .map((x) => (
                <p key={x.id}>
                  <SafetyCertificateOutlined />
                  {x.name}
                </p>
              ))}
          </div>
          <div className="summary-block">
            <span>
              输出字段 <Tag>{selectedFields.length}</Tag>
            </span>
            <div>
              {fields
                .filter((x) => selectedFields.includes(x.key))
                .slice(0, 9)
                .map((x) => (
                  <Tag key={x.key}>{x.label}</Tag>
                ))}
            </div>
          </div>
          <Button
            type="primary"
            size="large"
            block
            loading={loading}
            disabled={
              !selectedFields.length ||
              (workMode === "static" && !selectedObjects.length)
            }
            onClick={run}
          >
            {workMode === "static"
              ? "执行静态筛选"
              : product === canonicalProduct
                ? "运行动态评估"
                : "运行内外数融合组合"}
          </Button>
          <Button block onClick={savePlan}>
            保存研究方案
          </Button>
        </Card>
      </div>
      <Card
        className="wb-result wind-result"
        title={
          workMode === "static"
            ? "5. 静态筛选结果"
            : product === canonicalProduct
              ? "5. 企业动态机会与风险评估结果"
              : "5. 内外数融合结果"
        }
        extra={
          generated && (
            <Tag color="green">
              已生成{" "}
              {workMode === "static"
                ? selectedObjects.length
                : effectiveRows.length}{" "}
              条
            </Tag>
          )
        }
      >
        {!generated ? (
          <Empty
            description={
              workMode === "static"
                ? "添加研究对象、选择外数指标和筛选条件后运行，可按所选字段形成可导出数据表"
                : product === canonicalProduct
                  ? "运行后将展示三维状态、较上期变化、变化原因、置信度和建议动作"
                  : "选择内外数指标并运行后，将生成评分、判断依据与业务建议"
            }
          />
        ) : (
          <>
            <ProductInsightPanel
              product={workMode === "static" ? (selectedAnalysisProducts[0] || product) : product}
              rows={workMode === "static" ? selectedObjects.map((name, i) => ({...effectiveRows[i % Math.max(effectiveRows.length, 1)], name})) : effectiveRows}
              parameters={activeParameters}
            />
            {workMode === "product" && product === canonicalProduct && (
              <div className="dynamic-overview">
                <div>
                  <small>评估对象</small>
                  <b>{effectiveRows.length} 家</b>
                </div>
                <div>
                  <small>机会指数上升</small>
                  <b>
                    {
                      effectiveRows.filter(
                        (x: any) => Number(x.opportunityChange) > 0,
                      ).length
                    }{" "}
                    家
                  </b>
                </div>
                <div>
                  <small>重点拓展/优先营销</small>
                  <b>
                    {
                      effectiveRows.filter((x: any) =>
                        ["重点拓展", "优先营销"].includes(x.decision),
                      ).length
                    }{" "}
                    家
                  </b>
                </div>
                <div className="risk">
                  <small>风险升级处置</small>
                  <b>
                    {
                      effectiveRows.filter(
                        (x: any) => x.decision === "风险升级处置",
                      ).length
                    }{" "}
                    家
                  </b>
                </div>
              </div>
            )}
            <div className="result-summary">
              <span>
                本次结果{" "}
                <b>
                  {workMode === "static"
                    ? selectedObjects.length
                    : effectiveRows.length}
                </b>
              </span>
              <span>
                外数指标 <b>{selectedResources.length}</b>
              </span>
              <span>
                内数指标 <b>{selectedInternal.length}</b>
              </span>
              <span>
                {workMode === "static" ? "主体匹配与去重" : "动态状态已更新"}{" "}
                <b>已完成</b>
              </span>
            </div>
            <Table
              size="small"
              scroll={{ x: 1800, y: 390 }}
              pagination={{
                pageSize: 10,
                showSizeChanger: true,
                pageSizeOptions: [10, 20, 50, 100],
                showTotal: (total) => `共 ${total} 条`,
              }}
              dataSource={
                workMode === "static"
                  ? selectedObjects.map((name, i) => ({
                      ...effectiveRows[i % Math.max(effectiveRows.length, 1)],
                      name,
                    }))
                  : effectiveRows
              }
              rowKey={(r: any, i: number) => `${r.caseNo || r.name}-${i}`}
              columns={resultColumns}
            />
            <Space>
              <Button
                icon={<DownloadOutlined />}
                onClick={() =>
                  exportExcel(
                    workMode === "static"
                      ? selectedObjects.map((name, i) => ({
                          ...effectiveRows[i % Math.max(effectiveRows.length, 1)],
                          name,
                        }))
                      : effectiveRows,
                    runtime.file,
                    selectedFields,
                  )
                }
              >
                按所选字段导出Excel
              </Button>
              {workMode === "static" ? (
                <>
                  <Button
                    onClick={() =>
                      message.success(
                        "当前结果已保存为可复用名单，可被其他产品引用",
                      )
                    }
                  >
                    保存为名单
                  </Button>
                  <Button
                    type="primary"
                    onClick={() => {
                      setWorkMode("product");
                      message.success("静态名单已带入标准产品作为研究对象");
                    }}
                  >
                    带入标准产品持续分析
                  </Button>
                </>
              ) : (
                <>
                  <Button
                    type="primary"
                    disabled={pushed}
                    onClick={() => {
                      setPushed(true);
                      localStorage.setItem("icbc-feedback-product", product);
                      message.success(runtime.push);
                    }}
                  >
                    {pushed ? "已推送并生成任务编号" : runtime.push}
                  </Button>
                  <Button
                    onClick={() => {
                      localStorage.setItem("icbc-feedback-product", product);
                      go("feedback");
                    }}
                  >
                    评价产品
                  </Button>
                </>
              )}
            </Space>
          </>
        )}
      </Card>
      <Modal
        width={850}
        title="从外数资源库补充资源"
        open={resourceOpen}
        onCancel={() => setResourceOpen(false)}
        onOk={() => {
          setResourceOpen(false);
          message.success(`已保留 ${selectedResources.length} 项外数资源`);
        }}
        okText="确认添加"
      >
        <Input
          prefix={<SearchOutlined />}
          value={resourceQuery}
          onChange={(e) => setResourceQuery(e.target.value)}
          placeholder="搜索资源名称、编号、供应商或字段"
          allowClear
        />
        <div className="wb-resource-picker">
          {shownResources.map((r: any) => (
            <label
              key={r.id}
              className={selectedResources.includes(r.id) ? "selected" : ""}
            >
              <Checkbox
                checked={selectedResources.includes(r.id)}
                onChange={(e) =>
                  setSelectedResources((v) =>
                    e.target.checked
                      ? [...new Set([...v, r.id])]
                      : v.filter((x) => x !== r.id),
                  )
                }
              />
              <span>
                <b>{r.name}</b>
                <small>
                  {r.id} · {r.supplier} · {r.type}
                </small>
                <em>{r.fields}</em>
              </span>
              <Tag color={r.auth.includes("授权") ? "green" : "orange"}>
                {r.auth}
              </Tag>
            </label>
          ))}
        </div>
      </Modal>
    </div>
  );
}

function Feedback({ onReview }: { onReview: () => void }) {
  const [done, setDone] = useState(false);
  const [product, setProduct] = useState(() =>
    typeof window !== "undefined"
      ? localStorage.getItem("icbc-feedback-product") || "专精特新企业营销名单"
      : "专精特新企业营销名单",
  );
  const [issueType, setIssueType] = useState("新增需求");
  const [comment, setComment] = useState(
    "地方资质口径准确，名单已顺利进入营销任务池。建议后续增加客户触达结果的分支机构对比。",
  );
  const submitReview = () => {
    const record = {
      id: `FB-${Date.now()}`,
      product,
      type: issueType,
      text: comment,
      score: 5,
      time: new Date().toISOString(),
    };
    localStorage.setItem("icbc-review", "true");
    localStorage.setItem("icbc-review-record", JSON.stringify(record));
    setDone(true);
    onReview();
    message.success("评价已提交并生成待处理反馈");
  };
  if (done)
    return (
      <Card>
        <Result
          status="success"
          title="评价提交成功"
          subTitle="感谢反馈。评价已同步至产品运营评价，产品管理员可据此创建优化事项。"
          extra={
            <Button onClick={() => setDone(false)}>继续评价其他产品</Button>
          }
        />
      </Card>
    );
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            评价反馈 <DemoTag />
          </h1>
          <p>对产品业务价值、数据质量和使用体验进行结构化评价</p>
        </div>
      </div>
      <Row gutter={16}>
        <Col span={16}>
          <Card title="提交产品评价">
            <Form layout="vertical">
              <Form.Item label="评价产品">
                <Select
                  value={product}
                  onChange={setProduct}
                  options={Object.keys(workbenchProducts).map((value) => ({
                    value,
                  }))}
                />
              </Form.Item>
              <div className="rating-grid">
                {[
                  "业务匹配程度",
                  "数据准确程度",
                  "数据时效性",
                  "结果可理解程度",
                  "操作便利程度",
                ].map((x) => (
                  <div key={x}>
                    <span>{x}</span>
                    <Rate defaultValue={5} />
                  </div>
                ))}
              </div>
              <Form.Item label="问题类型">
                <Select
                  value={issueType}
                  onChange={setIssueType}
                  options={[
                    "数据错误",
                    "更新延迟",
                    "功能异常",
                    "操作困难",
                    "结果解释不足",
                    "新增需求",
                  ].map((value) => ({ value }))}
                />
              </Form.Item>
              <Form.Item label="评价与改进建议">
                <Input.TextArea
                  rows={5}
                  value={comment}
                  onChange={(e) => setComment(e.target.value)}
                />
              </Form.Item>
              <Button
                type="primary"
                size="large"
                disabled={!comment.trim()}
                onClick={submitReview}
              >
                提交评价
              </Button>
            </Form>
          </Card>
        </Col>
        <Col span={8}>
          <Card title="评价如何被使用">
            <Timeline
              items={[
                { children: "用户提交评分与改进建议" },
                { children: "进入管理门户运营评价" },
                { children: "管理员创建优化事项" },
                { children: "指定责任人与完成期限" },
                { children: "新版本发布后反馈用户" },
              ]}
            />
          </Card>
        </Col>
      </Row>
    </div>
  );
}

const sceneCatalog = [
  {
    name: "客户拓展",
    icon: <TeamOutlined />,
    tasks: ["目标客户发现", "招投标线索挖掘", "专精特新客群筛选"],
    products: [
      "专精特新企业营销名单",
      "招投标客户发现",
      "重点项目融资需求跟踪",
    ],
    resources: [
      "全国企业工商登记数据库",
      "全国招投标公告库",
      "专精特新企业名单",
    ],
    count: 18,
  },
  {
    name: "授信调查",
    icon: <FileDoneOutlined />,
    tasks: ["企业基本面核验", "集团关系识别", "外部风险排查"],
    products: [
      "授信尽调外部信息核验",
      "集团客户关联关系识别",
      "企业司法风险监测",
    ],
    resources: ["企业股权关系API", "司法涉诉风险查询", "行政许可与处罚信息库"],
    count: 14,
  },
  {
    name: "风险监测",
    icon: <SafetyCertificateOutlined />,
    tasks: ["司法风险监测", "舆情风险预警", "供应链风险传导"],
    products: ["企业司法风险监测", "舆情风险预警", "核心企业供应链风险画像"],
    resources: ["经营异常与失信API", "社会舆情聚合资讯", "供应链交易关系图谱"],
    count: 21,
  },
  {
    name: "产业研究",
    icon: <LineChartOutlined />,
    tasks: ["区域产业画像", "产业链分析", "景气趋势研判"],
    products: [
      "区域产业链客户画像",
      "区域产业景气监测",
      "产业事件驱动的融资机会识别",
    ],
    resources: ["产业链企业图谱", "区域产业景气指标库", "产业研究报告库"],
    count: 16,
  },
  {
    name: "经营分析",
    icon: <TableOutlined />,
    tasks: ["区域经营分析", "客群结构分析", "市场机会监测"],
    products: ["区域产业景气监测", "产业事件驱动的融资机会识别"],
    resources: [
      "宏观经济指标库",
      "大宗商品价格与库存",
      "企业信用评分与预警标签",
    ],
    count: 12,
  },
  {
    name: "监管与合规",
    icon: <AuditOutlined />,
    tasks: ["行政处罚监测", "监管政策跟踪", "客户合规核验"],
    products: ["行政处罚与合规事项监测", "授信尽调外部信息核验"],
    resources: ["行政许可与处罚信息库", "重点项目与产业政策库"],
    count: 9,
  },
];

const smartFindExamples = [
  { key:"marketing", icon:<TeamOutlined />, title:"找潜在客户", desc:"从产业资质、招投标和经营变化中发现融资机会", query:"查找四川省动力电池产业链中近一年有扩产或中标事件、近期无重大司法风险的潜在客户", scene:"客户拓展", task:"目标客户发现", parse:["动力电池产业链企业","四川省","扩产/中标","机会企业＋风险排除"] },
  { key:"credit", icon:<FileDoneOutlined />, title:"做授信尽调", desc:"一次核验工商、股权、司法、舆情与集团关系", query:"核验华北智造装备有限公司的工商股权、集团关系、司法风险和近半年舆情，并生成授信尽调摘要", scene:"授信调查", task:"企业基本面核验", parse:["华北智造装备","工商/股权/司法/舆情","近半年","尽调摘要＋证据"] },
  { key:"risk", icon:<SafetyCertificateOutlined />, title:"查风险变化", desc:"持续监测风险事件并关联行内敞口和处置任务", query:"监测存量制造业客户近30日司法、经营异常和负面舆情变化，优先提示有授信敞口的企业", scene:"风险监测", task:"司法风险监测", parse:["存量制造业客户","司法/经营异常/舆情","近30日","风险分级＋处置任务"] },
  { key:"industry", icon:<LineChartOutlined />, title:"看产业机会", desc:"分析产业景气、链上影响、重点区域和企业行动", query:"分析长三角高端装备产业链近期景气变化、机会风险传导和重点企业，形成可导出的经营简报", scene:"产业研究", task:"产业链分析", parse:["高端装备产业链","长三角","近期景气与事件","热力图＋传导图＋企业明细"] },
];

const matchSmartDemand = (text: string) => {
  const q = text.trim();
  let sceneName = "客户拓展";
  if (/授信|尽调|股权|受益人/.test(q)) sceneName = "授信调查";
  else if (/风险|司法|舆情|失信|处罚|预警/.test(q)) sceneName = "风险监测";
  else if (/产业链|景气|产业研究|行业/.test(q)) sceneName = "产业研究";
  else if (/经营分析|客群结构|市场机会/.test(q)) sceneName = "经营分析";
  else if (/监管|合规|行政处罚/.test(q)) sceneName = "监管与合规";
  const base = sceneCatalog.find((x) => x.name === sceneName) || sceneCatalog[0];
  const task =
    base.tasks.find((x) =>
      (x.includes("司法") && /司法|失信|诉讼/.test(q)) ||
      (x.includes("舆情") && /舆情|声誉/.test(q)) ||
      (x.includes("产业链") && /产业链|上下游/.test(q)) ||
      (x.includes("招投标") && /招标|中标|项目/.test(q)) ||
      (x.includes("基本面") && /工商|股权|尽调/.test(q)),
    ) || base.tasks[0];
  const region = q.match(/(全国|北京市|上海市|广东省|浙江省|江苏省|四川省|长三角|珠三角|京津冀)/)?.[0] || "未限定地域";
  const windowText = q.match(/(近\d+[日月年]|近半年|近一年|实时|持续)/)?.[0] || "按最新可用数据";
  const object = q.match(/([\u4e00-\u9fa5A-Za-z0-9]{2,24}(?:有限公司|股份有限公司|企业|客户|产业链))/)?.[0] || (sceneName === "产业研究" ? "目标产业链及链上企业" : "目标企业/客户");
  const delivery = /简报|摘要|报告/.test(q) ? "分析摘要＋事实证据" : /名单|潜在客户/.test(q) ? "企业名单＋推荐依据" : /预警|风险/.test(q) ? "风险分级＋处置建议" : "画像＋事实证据＋行动建议";
  let products = [...base.products];
  let resources = [...base.resources];
  if (/动力电池|新能源|扩产|融资机会/.test(q)) {
    products = ["全国产业链机会与风险识别", "区域产业链客户画像", "招投标客户发现"];
    resources = ["产业链企业图谱", "全国企业工商登记数据库", "重点项目与产业政策库", "全国招投标公告库", "司法涉诉风险查询"];
  } else if (/专精特新/.test(q)) {
    products = ["全行专精特新企业营销名单", "招投标客户发现", "企业司法风险监测"];
    resources = ["专精特新企业名单", "全国企业工商登记数据库", "全国招投标公告库", "司法涉诉风险查询"];
  } else if (/股权|集团关系|受益人/.test(q)) {
    products = ["授信尽调外部信息核验", "集团客户关联关系识别", "企业司法风险监测"];
    resources = ["企业股权关系API", "全国企业工商登记数据库", "司法涉诉风险查询", "行政许可与处罚信息库"];
  }
  return { scene: { ...base, products, resources }, task, parse: [object, region, windowText, delivery] };
};

function SceneFinder({ go }: { go: (v: View) => void }) {
  const [scene, setScene] = useState(sceneCatalog[0]);
  const [task, setTask] = useState(sceneCatalog[0].tasks[0]);
  const [activeExample, setActiveExample] = useState(smartFindExamples[0]);
  const [query, setQuery] = useState(smartFindExamples[0].query);
  const [analyzed, setAnalyzed] = useState(false);
  const [thinking, setThinking] = useState(false);
  const [parsedDemand, setParsedDemand] = useState(smartFindExamples[0].parse);
  const choose = (x: any) => {
    setScene(x);
    setTask(x.tasks[0]);
  };
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            智能找数 <DemoTag />
          </h1>
          <p>直接说清楚你要找谁、判断什么、希望得到什么，平台自动组合可用产品与外数资源。</p>
        </div>
      </div>
      <div className="smart-example-grid">
        {smartFindExamples.map(x=><button key={x.key} className={activeExample.key===x.key?"active":""} onClick={()=>{setActiveExample(x);setQuery(x.query);setParsedDemand(x.parse);setAnalyzed(false);setThinking(false);const target=sceneCatalog.find(s=>s.name===x.scene)||sceneCatalog[0];setScene(target);setTask(x.task)}}><span>{x.icon}</span><div><b>{x.title}</b><small>{x.desc}</small></div><em>填入示例</em></button>)}
      </div>
      <Card className="smart-find-card smart-find-business">
        <div className="smart-prompt-title"><div><RobotOutlined /><span><b>今天想完成什么业务任务？</b><small>可以直接输入一句话，也可以点击上方示例开始</small></span></div><Tag color="green">可用资源已按权限过滤</Tag></div>
        <Input.TextArea value={query} onChange={(e) => {setQuery(e.target.value);setAnalyzed(false);setThinking(false)}} autoSize={{ minRows: 3, maxRows: 5 }} placeholder="例如：帮我查找北京近三年成立、有招投标活动且没有重大司法风险的专精特新制造企业" />
        <div className="smart-prompt-actions"><Space wrap><span>常用条件：</span>{["限定地区","限定行业","近一年事件","排除重大风险","生成经营简报"].map(x=><Tag key={x} onClick={()=>{setQuery(v=>`${v}${v?"，":""}${x}`);setAnalyzed(false);setThinking(false)}}>{x}</Tag>)}</Space><Space><Button onClick={()=>{setQuery("");setAnalyzed(false);setThinking(false)}}>重新输入</Button><Button type="primary" size="large" icon={<SearchOutlined />} loading={thinking} disabled={!query.trim()} onClick={() => { const match=matchSmartDemand(query); setScene(match.scene); setTask(match.task); setParsedDemand(match.parse); setAnalyzed(false); setThinking(true); window.setTimeout(()=>{setThinking(false);setAnalyzed(true)},1500); }}>帮我找数</Button></Space></div>
        {analyzed && <div className="smart-understanding"><b>我理解你的需求是：</b>{parsedDemand.map((x,i)=><React.Fragment key={`${x}-${i}`}><span>{["查询对象","范围与条件","观察期限","希望结果"][i]}：<strong>{x}</strong></span>{i<3&&<em>→</em>}</React.Fragment>)}<Button type="link" onClick={()=>setAnalyzed(false)}>调整理解</Button></div>}
      </Card>
      {thinking ? (
        <Card className="smart-ai-thinking"><div className="ai-thinking-head"><RobotOutlined /><div><b>正在理解你的业务需求</b><small>我会先判断业务场景和交付目标，再核对权限、时效与可用资源，最后给出可解释的组合建议。</small></div></div><Steps size="small" current={1} items={[{title:"理解任务"},{title:"检索产品与资源"},{title:"比较匹配度"},{title:"形成建议"}]} /></Card>
      ) : !analyzed ? (
        <Card className="smart-waiting"><Empty image={Empty.PRESENTED_IMAGE_SIMPLE} description="示例只会填入业务问题；点击“帮我找数”后，平台才会开始分析并生成推荐。" /></Card>
      ) : <><Card className="smart-ai-answer"><div className="ai-answer-title"><Avatar icon={<RobotOutlined />} /><div><b>为你推荐：优先调用“{scene.products[0]}”</b><small>以下建议综合考虑任务匹配度、数据授权、更新频率和结果可直接使用程度。</small></div><Tag color="green">推荐置信度 92%</Tag></div><p>你的目标不是单独查询某一项数据，而是完成“{task}”。因此，建议以标准产品作为主入口，组合调用{scene.resources.join("、")}，由平台完成主体关联、指标计算和规则判断；这样可以直接获得可解释的业务结果，并保留来源证据。</p><div className="ai-reason-grid"><span><b>为什么推荐</b><small>业务任务与产品流程高度一致，可减少手工拼接</small></span><span><b>你将得到</b><small>企业画像、事实证据、风险提示与行动建议</small></span><span><b>使用前注意</b><small>已按当前岗位权限过滤，地方口径可继续调整</small></span></div></Card><div className="scene-layout">
        <Card className="scene-nav-card" title="1. 语义匹配场景">
          {sceneCatalog.map((x) => (
            <button
              key={x.name}
              className={scene.name === x.name ? "active" : ""}
              onClick={() => choose(x)}
            >
              <span>{x.icon}</span>
              <b>{x.name}</b>
              <small>{x.count} 个可用成果</small>
            </button>
          ))}
        </Card>
        <div>
          <Card title="2. 校准识别结果" className="task-card">
            <Space wrap>
              {scene.tasks.map((x) => (
                <Button
                  key={x}
                  type={task === x ? "primary" : "default"}
                  onClick={() => setTask(x)}
                >
                  {x}
                </Button>
              ))}
            </Space>
            <div className="task-path">
              <span>{scene.name}</span>
              <em>→</em>
              <span>{task}</span>
              <em>→</em>
              <b>产品与资源组合推荐</b>
            </div>
          </Card>
          <Card
            title="3. 产品与资源推荐"
            extra={<Tag color="green">匹配度 92%</Tag>}
            className="recommendation-card"
          >
            <div className="recommendation-head">
              <div>
                <span className="product-icon">
                  <SolutionOutlined />
                </span>
                <div>
                  <b>{scene.products[0]}</b>
                  <p>
                    已将业务任务、输入数据、加工能力与结果交付组合为可直接使用的标准产品。
                  </p>
                </div>
              </div>
              <Button
                type="primary"
                onClick={() => {
                  const target = workbenchProducts[scene.products[0]]
                    ? scene.products[0]
                    : scene.name === "风险监测"
                      ? "企业司法风险监测"
                      : "专精特新企业营销名单";
                  localStorage.setItem("icbc-workbench-product", target);
                  go("workbench");
                }}
              >
                带入工作台
              </Button>
            </div>
            <div className="bundle-flow">
              <div>
                <small>业务任务</small>
                <b>{task}</b>
              </div>
              <em>＋</em>
              <div>
                <small>外数资源</small>
                <b>{scene.resources.length} 项</b>
              </div>
              <em>＋</em>
              <div>
                <small>能力模块</small>
                <b>4 项</b>
              </div>
              <em>＋</em>
              <div>
                <small>通用规则</small>
                <b>6 条</b>
              </div>
              <em>→</em>
              <div className="result">
                <small>结果交付</small>
                <b>名单／画像／提醒</b>
              </div>
            </div>
            <Row gutter={14}>
              <Col span={12}>
                <h3>推荐产品</h3>
                <List
                  size="small"
                  dataSource={scene.products}
                  renderItem={(x, i) => (
                    <List.Item
                      actions={[
                        <Button
                          key="use"
                          size="small"
                          type={i === 0 ? "primary" : "default"}
                          onClick={() =>
                            go(i === 0 ? "product-detail" : "app-products")
                          }
                        >
                          {i === 0 ? "立即使用" : "查看"}
                        </Button>,
                      ]}
                    >
                      <List.Item.Meta
                        avatar={<Avatar className="scene-rank">{i + 1}</Avatar>}
                        title={x}
                        description={`${[92, 86, 78][i]}% 任务匹配 · ${i === 0 ? "全行标准产品" : "相关产品"}`}
                      />
                    </List.Item>
                  )}
                />
              </Col>
              <Col span={12}>
                <h3>匹配的外数资源</h3>
                <List
                  size="small"
                  dataSource={scene.resources}
                  renderItem={(x, i) => (
                    <List.Item>
                      <DatabaseOutlined className="red" />
                      <span className="scene-resource"><b>{x}</b><small>{i < 2 ? "完成任务的核心数据" : "用于交叉验证与补充判断"}</small></span>
                      <Tag color={i < 2 ? "red" : "default"}>{i < 2 ? "核心必需" : "补充验证"}</Tag>
                    </List.Item>
                  )}
                />
              </Col>
            </Row>
          </Card>
        </div>
      </div></>}
    </div>
  );
}

function RelationView({ go }: { go: (v: View) => void }) {
  const [focus, setFocus] = useState("专精特新企业营销名单");
  const relationSets: any = {
    专精特新企业营销名单: {
      resources: [
        "工商登记数据库",
        "专精特新名单",
        "招投标公告库",
        "司法涉诉查询",
      ],
      abilities: ["主体关联", "产业资质识别", "名单筛选", "司法风险核验"],
      rules: ["地区筛选", "注册年限", "风险排除"],
      scene: "客户拓展",
    },
    企业司法风险监测: {
      resources: ["司法涉诉查询", "经营异常与失信", "社会舆情资讯"],
      abilities: ["主体关联", "司法风险核验", "舆情事件识别"],
      rules: ["司法风险分级", "经营异常排除"],
      scene: "风险监测",
    },
    区域产业链客户画像: {
      resources: ["产业链企业图谱", "工商登记数据库", "产业研究报告库"],
      abilities: ["主体关联", "企业画像", "指标计算"],
      rules: ["地区筛选", "行业筛选"],
      scene: "产业研究",
    },
  };
  const x = relationSets[focus];
  const nodes = [
    { name: focus, category: 0, symbolSize: 72 },
    { name: x.scene, category: 1, symbolSize: 48 },
    ...x.resources.map((name: string) => ({
      name,
      category: 2,
      symbolSize: 40,
    })),
    ...x.abilities.map((name: string) => ({
      name,
      category: 3,
      symbolSize: 40,
    })),
    ...x.rules.map((name: string) => ({ name, category: 4, symbolSize: 34 })),
  ];
  const links = nodes
    .slice(1)
    .map((n: any) => ({ source: focus, target: n.name }));
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            关联视图 <DemoTag />
          </h1>
          <p>
            从产品出发查看业务场景、外数资源、能力模块和通用规则的完整调用关系。
          </p>
        </div>
        <Space>
          <Select
            value={focus}
            style={{ width: 260 }}
            onChange={setFocus}
            options={Object.keys(relationSets).map((value) => ({ value }))}
          />
          <Button onClick={() => message.success("关系图已导出")}>
            导出关系图
          </Button>
        </Space>
      </div>
      <Row gutter={14}>
        <Col span={18}>
          <Card
            title="产品关联网络"
            extra={
              <Space>
                <Tag color="red">产品</Tag>
                <Tag color="blue">场景</Tag>
                <Tag color="green">资源</Tag>
                <Tag color="orange">能力</Tag>
                <Tag color="purple">规则</Tag>
              </Space>
            }
          >
            <ReactECharts
              style={{ height: 540 }}
              option={{
                tooltip: {},
                legend: { show: false },
                series: [
                  {
                    type: "graph",
                    layout: "force",
                    roam: true,
                    label: { show: true, fontSize: 11 },
                    force: { repulsion: 360, edgeLength: [90, 165] },
                    categories: [
                      { itemStyle: { color: "#c7000b" } },
                      { itemStyle: { color: "#316bad" } },
                      { itemStyle: { color: "#28966f" } },
                      { itemStyle: { color: "#d77c00" } },
                      { itemStyle: { color: "#7651a8" } },
                    ],
                    data: nodes,
                    links,
                    lineStyle: {
                      color: "#aeb6c2",
                      width: 1.5,
                      curveness: 0.08,
                    },
                  },
                ],
              }}
            />
          </Card>
        </Col>
        <Col span={6}>
          <Card title="关系摘要" className="relation-summary">
            <Statistic
              title="关联外数资源"
              value={x.resources.length}
              suffix="项"
            />
            <Statistic
              title="复用能力模块"
              value={x.abilities.length}
              suffix="项"
            />
            <Statistic
              title="调用通用规则"
              value={x.rules.length}
              suffix="条"
            />
            <Progress
              type="circle"
              percent={82}
              strokeColor="#c7000b"
              size={110}
            />
            <p>产品资产复用率</p>
            <Button block type="primary" onClick={() => go("product-detail")}>
              查看产品详情
            </Button>
          </Card>
          <Card title="上下游影响">
            <Timeline
              items={[
                { children: "资源更新触发产品重新计算" },
                { children: "能力版本升级影响 3 个产品" },
                { children: "用户评价进入运营优化" },
                { children: "结果推送客户经理任务池" },
              ]}
            />
          </Card>
        </Col>
      </Row>
    </div>
  );
}

function FollowCenter({ go }: { go: (v: View) => void }) {
  const rawFollows =
    typeof window === "undefined"
      ? null
      : localStorage.getItem("icbc-follow-products");
  const stored = rawFollows ? JSON.parse(rawFollows) : defaultFollowProducts;
  const storedRes =
    typeof window === "undefined"
      ? []
      : JSON.parse(localStorage.getItem("icbc-follow-resources") || "[]");
  const [productNames, setProductNames] = useState<string[]>(
    Array.from(new Set(stored)),
  );
  const [resourceIds, setResourceIds] = useState<string[]>(
    Array.from(new Set(["EXT-R003", "EXT-R004", "EXT-R007", ...storedRes])),
  );
  const [personalPlans, setPersonalPlans] = useState<any[]>(() =>
    typeof window === "undefined"
      ? []
      : JSON.parse(localStorage.getItem("icbc-personal-plans") || "[]"),
  );
  useEffect(() => {
    const p = JSON.parse(localStorage.getItem("icbc-personal-plans") || "[]");
    const raw = localStorage.getItem("icbc-follow-products");
    const fp = raw ? JSON.parse(raw) : defaultFollowProducts;
    const fr = JSON.parse(
      localStorage.getItem("icbc-follow-resources") || "[]",
    );
    setPersonalPlans(p);
    setProductNames(Array.from(new Set(fp)));
    setResourceIds(
      Array.from(new Set(["EXT-R003", "EXT-R004", "EXT-R007", ...fr])),
    );
  }, []);
  const removeProduct = (name: string) => {
    const next = productNames.filter((x) => x !== name);
    setProductNames(next);
    localStorage.setItem("icbc-follow-products", JSON.stringify(next));
    message.success("已取消关注");
  };
  const removeResource = (id: string) => {
    setResourceIds((v) => v.filter((x) => x !== id));
    message.success("已移出自选");
  };
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            我的关注与自选 <DemoTag />
          </h1>
          <p>集中管理关注产品、自选资源、业务主题和更新提醒。</p>
        </div>
        <Space>
          <Button onClick={() => go("app-resources")}>添加资源</Button>
          <Button type="primary" onClick={() => go("app-products")}>
            关注产品
          </Button>
        </Space>
      </div>
      <Row gutter={14} className="follow-metrics">
        <Col span={6}>
          <MetricCard
            title="关注产品"
            value={productNames.length}
            trend="1 项今日更新"
          />
        </Col>
        <Col span={6}>
          <MetricCard
            title="自选资源"
            value={resourceIds.length}
            trend="2 项字段变更"
          />
        </Col>
        <Col span={6}>
          <MetricCard title="关注主题" value={4} trend="6 条新动态" />
        </Col>
        <Col span={6}>
          <MetricCard title="风险提醒" value={3} trend="1 条高优先级" alert />
        </Col>
      </Row>
      <Card>
        <Tabs
          items={[
            {
              key: "products",
              label: `关注产品（${productNames.length}）`,
              children: (
                <Table
                  pagination={false}
                  rowKey="name"
                  dataSource={productNames.map((name, i) => ({
                    name,
                    category: ["客户营销", "风险监测", "产业研究"][i % 3],
                    update: [
                      "V3.2版本发布",
                      "新增司法事件原文入口",
                      "长三角产业链数据更新",
                    ][i % 3],
                    time: ["今天 09:20", "昨天 16:40", "07-29 11:10"][i % 3],
                  }))}
                  columns={[
                    {
                      title: "产品名称",
                      dataIndex: "name",
                      render: (x: string) => (
                        <a
                          onClick={() => {
                            localStorage.setItem("icbc-selected-product", x);
                            go("product-detail");
                          }}
                        >
                          {x}
                        </a>
                      ),
                    },
                    {
                      title: "业务分类",
                      dataIndex: "category",
                      render: (x: string) => <Tag>{x}</Tag>,
                    },
                    { title: "最新动态", dataIndex: "update" },
                    { title: "更新时间", dataIndex: "time" },
                    { title: "提醒", render: () => <SwitchCell /> },
                    {
                      title: "操作",
                      render: (_: any, r: any) => (
                        <Space>
                          <Button
                            size="small"
                            onClick={() => {
                              localStorage.setItem(
                                "icbc-selected-product",
                                r.name,
                              );
                              go("product-detail");
                            }}
                          >
                            查看
                          </Button>
                          <Button
                            size="small"
                            danger
                            onClick={() => removeProduct(r.name)}
                          >
                            取消关注
                          </Button>
                        </Space>
                      ),
                    },
                  ]}
                />
              ),
            },
            {
              key: "resources",
              label: `自选资源（${resourceIds.length}）`,
              children: (
                <Table
                  pagination={false}
                  rowKey="id"
                  dataSource={resources.filter((r) =>
                    resourceIds.includes(r.id),
                  )}
                  columns={[
                    {
                      title: "资源名称",
                      dataIndex: "name",
                      render: (x: string) => (
                        <a onClick={() => go("app-resources")}>{x}</a>
                      ),
                    },
                    {
                      title: "类型",
                      dataIndex: "type",
                      render: (x: string) => <Tag>{x}</Tag>,
                    },
                    { title: "供应商", dataIndex: "supplier" },
                    { title: "更新频率", dataIndex: "frequency" },
                    {
                      title: "授权状态",
                      dataIndex: "auth",
                      render: (x: string) => <StatusTag status={x} />,
                    },
                    {
                      title: "操作",
                      render: (_: any, r: any) => (
                        <Button
                          size="small"
                          danger
                          onClick={() => removeResource(r.id)}
                        >
                          移出自选
                        </Button>
                      ),
                    },
                  ]}
                />
              ),
            },
            {
              key: "plans",
              label: `个人方案（${personalPlans.length}）`,
              children: personalPlans.length ? (
                <Table
                  pagination={false}
                  rowKey="id"
                  dataSource={personalPlans}
                  columns={[
                    {
                      title: "方案名称",
                      dataIndex: "name",
                      render: (x: string) => <b>{x}</b>,
                    },
                    { title: "基于正式产品", dataIndex: "product" },
                    {
                      title: "补充后资源",
                      dataIndex: "resources",
                      render: (x: string[]) => <Tag>{x.length}项</Tag>,
                    },
                    {
                      title: "业务条件",
                      dataIndex: "conditions",
                      render: (x: any) => (
                        <span className="plan-condition-summary">
                          {Object.entries(x)
                            .slice(0, 3)
                            .map(([k, v]) => `${k}：${v}`)
                            .join("；")}
                        </span>
                      ),
                    },
                    { title: "保存时间", dataIndex: "updated" },
                    {
                      title: "操作",
                      render: (_: any, r: any) => (
                        <Space>
                          <Button
                            size="small"
                            type="primary"
                            onClick={() => {
                              localStorage.setItem(
                                "icbc-active-plan",
                                JSON.stringify(r),
                              );
                              go("workbench");
                              message.success("已调取个人方案");
                            }}
                          >
                            调取使用
                          </Button>
                          <Button
                            size="small"
                            danger
                            onClick={() => {
                              const next = personalPlans.filter(
                                (x) => x.id !== r.id,
                              );
                              setPersonalPlans(next);
                              localStorage.setItem(
                                "icbc-personal-plans",
                                JSON.stringify(next),
                              );
                            }}
                          >
                            删除
                          </Button>
                        </Space>
                      ),
                    },
                  ]}
                />
              ) : (
                <Empty description="尚未保存个人方案">
                  <Button type="primary" onClick={() => go("workbench")}>
                    前往产品工作台配置
                  </Button>
                </Empty>
              ),
            },
            {
              key: "topics",
              label: "关注主题（4）",
              children: (
                <List
                  dataSource={[
                    [
                      "长三角高端装备产业链",
                      "产业链新增 26 家企业，3 家核心企业发生工商变更",
                    ],
                    ["浙江专精特新企业", "第七批名单更新，新增 142 家企业"],
                    [
                      "存量客户司法风险",
                      "12 家客户风险状态变化，其中 1 家为高风险",
                    ],
                    ["新能源行业景气", "上游材料价格指数连续三周回升"],
                  ]}
                  renderItem={(x, i) => (
                    <List.Item
                      actions={[
                        <Button
                          key="view"
                          size="small"
                          onClick={() =>
                            go(i === 2 ? "app-products" : "scenes")
                          }
                        >
                          查看动态
                        </Button>,
                      ]}
                    >
                      <List.Item.Meta
                        avatar={
                          <Badge count={i + 2}>
                            <Avatar icon={<NotificationOutlined />} />
                          </Badge>
                        }
                        title={x[0]}
                        description={x[1]}
                      />
                    </List.Item>
                  )}
                />
              ),
            },
          ]}
        />
      </Card>
    </div>
  );
}
function SwitchCell() {
  const [on, setOn] = useState(true);
  return (
    <Checkbox checked={on} onChange={(e) => setOn(e.target.checked)}>
      接收
    </Checkbox>
  );
}

const supplierData = [
  {
    name: "启明数据",
    domain: "企业数据",
    service: "API服务",
    status: "在用",
    resources: 12,
    products: 18,
    quality: "A",
    api: "99.96%",
    desc: "提供企业工商、股权关系与主体核验服务",
  },
  {
    name: "法信数据",
    domain: "企业风险",
    service: "查询服务",
    status: "在用",
    resources: 8,
    products: 14,
    quality: "A",
    api: "99.91%",
    desc: "提供司法涉诉、执行与风险事件数据",
  },
  {
    name: "筑信数据",
    domain: "企业行为",
    service: "数据库",
    status: "在用",
    resources: 6,
    products: 9,
    quality: "B",
    api: "99.78%",
    desc: "提供全国招投标公告及企业活跃度数据",
  },
  {
    name: "新华财经",
    domain: "公共数据",
    service: "资讯研报",
    status: "在用",
    resources: 9,
    products: 11,
    quality: "B",
    api: "99.83%",
    desc: "提供资讯、政策和舆情聚合服务",
  },
  {
    name: "产业智图",
    domain: "行业数据",
    service: "图谱服务",
    status: "试用",
    resources: 4,
    products: 6,
    quality: "B",
    api: "99.62%",
    desc: "提供产业链节点和企业链属关系",
  },
  {
    name: "联合信用数据",
    domain: "企业分析",
    service: "模型评分",
    status: "准入中",
    resources: 3,
    products: 5,
    quality: "A",
    api: "—",
    desc: "提供信用评分、预警标签与解释信息",
  },
];

function SupplierCenter() {
  const [domain, setDomain] = useState("全部领域");
  const [selected, setSelected] = useState<any>(null);
  const shown = supplierData.filter(
    (x) => domain === "全部领域" || x.domain === domain,
  );
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            数据供应商 <DemoTag />
          </h1>
          <p>从供应商查看其可提供的数据服务、行内采购资源及所支撑产品。</p>
        </div>
        <Input
          prefix={<SearchOutlined />}
          placeholder="搜索供应商或服务资源"
          style={{ width: 320 }}
        />
      </div>
      <div className="supplier-filter">
        <Select
          value={domain}
          onChange={setDomain}
          options={[
            "全部领域",
            "企业数据",
            "企业风险",
            "企业行为",
            "公共数据",
            "行业数据",
            "企业分析",
          ].map((value) => ({ value }))}
        />
        <Select
          defaultValue="全部合作状态"
          options={["全部合作状态", "在用", "试用", "准入中", "历史合作"].map(
            (value) => ({ value }),
          )}
        />
        <Select
          defaultValue="全部服务类型"
          options={[
            "全部服务类型",
            "API服务",
            "数据库",
            "查询服务",
            "资讯研报",
            "图谱服务",
          ].map((value) => ({ value }))}
        />
      </div>
      <div className="supplier-grid">
        {shown.map((x) => (
          <Card hoverable key={x.name} onClick={() => setSelected(x)}>
            <div className="supplier-head">
              <Avatar size={48} icon={<ShopOutlined />} />
              <div>
                <Space>
                  <Tag color="red">{x.domain}</Tag>
                  <StatusTag status={x.status} />
                </Space>
                <h2>{x.name}</h2>
              </div>
            </div>
            <p>{x.desc}</p>
            <div className="supplier-stats">
              <span>
                <b>{x.resources}</b>项资源
              </span>
              <span>
                <b>{x.products}</b>个产品
              </span>
              <span>
                <b>{x.quality}</b>质量等级
              </span>
            </div>
            <footer>
              <span>接口可用率 {x.api}</span>
              <b>查看供应商档案 →</b>
            </footer>
          </Card>
        ))}
      </div>
      <Drawer
        width={720}
        open={!!selected}
        onClose={() => setSelected(null)}
        title="供应商档案"
      >
        {selected && (
          <>
            <div className="supplier-detail">
              <Avatar size={58} icon={<ShopOutlined />} />
              <div>
                <h2>{selected.name}</h2>
                <p>{selected.desc}</p>
              </div>
            </div>
            <Descriptions
              bordered
              column={2}
              items={[
                {
                  label: "合作状态",
                  children: <StatusTag status={selected.status} />,
                },
                { label: "服务等级", children: "重要供应商" },
                { label: "覆盖领域", children: selected.domain },
                { label: "接口可用率", children: selected.api },
                { label: "已采购资源", children: selected.resources + " 项" },
                { label: "支撑产品", children: selected.products + " 个" },
              ]}
            />
            <h3>已采购数据与服务</h3>
            <Table
              size="small"
              pagination={false}
              dataSource={resources
                .filter((r) => r.supplier.includes(selected.name.slice(0, 2)))
                .slice(0, 4)}
              rowKey="id"
              columns={[
                { title: "资源名称", dataIndex: "name" },
                { title: "交付方式", dataIndex: "type" },
                { title: "更新频率", dataIndex: "frequency" },
                {
                  title: "授权",
                  dataIndex: "auth",
                  render: (x: string) => <StatusTag status={x} />,
                },
              ]}
            />
            <h3>服务评价</h3>
            <Row gutter={12}>
              {[
                ["数据质量", 4.8],
                ["更新时效", 4.7],
                ["服务响应", 4.6],
                ["成本效益", 4.4],
              ].map((x) => (
                <Col span={6} key={x[0]}>
                  <Card size="small">
                    <small>{x[0]}</small>
                    <Rate disabled allowHalf value={Number(x[1])} />
                    <b>{x[1]}</b>
                  </Card>
                </Col>
              ))}
            </Row>
          </>
        )}
      </Drawer>
    </div>
  );
}

function DataDemand({ go }: { go: (v: View) => void }) {
  const [form] = Form.useForm();
  const [submitted, setSubmitted] = useState(false);
  const [category, setCategory] = useState("客户营销");
  const [matched, setMatched] = useState(true);
  const [confirmedGap, setConfirmedGap] = useState(false);
  const recommendations =
    category === "风险监测"
      ? [
          {
            id: "P-2026-003",
            name: "企业司法风险监测",
            score: 94,
            scope: "全行",
            fit: "可直接使用",
            reason: "覆盖企业主体、司法风险监测、事件触发更新与全国地域。",
          },
          {
            id: "P-GD-2026-012",
            name: "核心企业供应链风险监测",
            score: 76,
            scope: "广东省",
            fit: "需地域适配",
            reason: "业务任务相近，但供应链关系范围及适用地域需调整。",
          },
        ]
      : category === "客户营销"
        ? [
            {
              id: "P-2026-002",
              name: "专精特新企业营销名单",
              score: 92,
              scope: "全行",
              fit: "可直接使用",
              reason:
                "业务任务、数据对象、重点字段和更新频率高度一致，地域可作为运行参数配置。",
            },
            {
              id: "P-2026-005",
              name: "招投标客户发现",
              score: 84,
              scope: "全行",
              fit: "组合使用",
              reason: "可补充企业招中标活跃度与项目融资线索。",
            },
            {
              id: "P-ZJ-2026-001",
              name: "浙江分行专精特新企业筛选产品",
              score: 71,
              scope: "浙江省",
              fit: "参考建设",
              reason: "任务链路一致，但地方名单来源和覆盖地域不同。",
            },
          ]
        : [
            {
              id: "P-2026-004",
              name: "区域产业链客户画像",
              score: 81,
              scope: "全行",
              fit: "需参数适配",
              reason: "可复用主体识别、行业映射与画像能力，具体输出仍需确认。",
            },
          ];
  const inspect = () => {
    form
      .validateFields()
      .then(() => {
        setMatched(true);
        setConfirmedGap(false);
        message.success(
          `已按业务任务、对象、地域、频率和重点字段匹配到 ${recommendations.length} 个类似产品`,
        );
      })
      .catch(() => {});
  };
  const useProduct = (name: string) => {
    localStorage.setItem("icbc-workbench-product", name);
    message.success(`已选择“${name}”，正在带入产品工作台`);
    go("workbench");
  };
  if (submitted)
    return (
      <Result
        status="success"
        title="未满足部分已作为数据需求上交"
        subTitle="需求编号 DR-2026-0804-021 · 已关联相似产品及差异说明，管理人员将优先复用现有产品、能力和资源补齐缺口。"
        extra={
          <Space>
            <Button
              onClick={() => {
                setSubmitted(false);
                setConfirmedGap(false);
              }}
            >
              继续提交需求
            </Button>
            <Button type="primary" onClick={() => go("follows")}>
              查看我的需求
            </Button>
          </Space>
        }
      />
    );
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            数据需求 <DemoTag />
          </h1>
          <p>
            先按统一业务任务描述需求，系统优先推荐可复用产品；现有产品无法满足时，再将差异部分上交。
          </p>
        </div>
        <Button>查看我的需求</Button>
      </div>
      <div className="demand-steps">
        <div className="active">
          <b>1</b>
          <span>
            描述业务需求<small>统一任务与数据口径</small>
          </span>
        </div>
        <div className={matched ? "active" : ""}>
          <b>2</b>
          <span>
            库内智能匹配<small>优先推荐类似产品</small>
          </span>
        </div>
        <div className={confirmedGap ? "active" : ""}>
          <b>3</b>
          <span>
            确认缺口并上交<small>仅提交未满足部分</small>
          </span>
        </div>
      </div>
      <Row gutter={14}>
        <Col span={15}>
          <Card
            title="业务数据需求描述"
            extra={<Tag color="red">与外数产品使用同一任务口径</Tag>}
          >
            <Form
              form={form}
              layout="vertical"
              initialValues={{
                category: "客户营销",
                task: "专精特新客群筛选",
                deadline: "2026-08-20",
                problem:
                  "筛选北京市近三年注册、制造业、国资控股、具有专精特新资质、近一年有招投标记录且近三个月无重大司法风险的企业。",
                objects: ["企业主体"],
                regions: ["北京市"],
                frequency: "每日",
                fields: [
                  "注册时间",
                  "行业",
                  "控股类型",
                  "专精特新资质",
                  "招投标记录",
                  "司法风险",
                ],
              }}
              onValuesChange={() => {
                setMatched(false);
                setConfirmedGap(false);
              }}
            >
              <Row gutter={14}>
                <Col span={12}>
                  <Form.Item
                    label="业务任务分类"
                    name="category"
                    rules={[{ required: true }]}
                  >
                    <Select
                      options={productTaskCategories.map((value) => ({
                        value,
                      }))}
                      onChange={(value) => {
                        setCategory(value);
                        form.setFieldValue("task", demandTaskOptions[value][0]);
                      }}
                    />
                  </Form.Item>
                </Col>
                <Col span={12}>
                  <Form.Item
                    label="具体业务任务"
                    name="task"
                    extra="可从建议任务中选择，也可直接输入本次需要完成的业务任务"
                    rules={[
                      { required: true, message: "请选择或填写具体业务任务" },
                    ]}
                  >
                    <AutoComplete
                      options={demandTaskOptions[category].map((value) => ({
                        value,
                      }))}
                      placeholder="请选择或自行填写具体业务任务"
                      filterOption={(input, option) =>
                        String(option?.value || "")
                          .toLowerCase()
                          .includes(input.toLowerCase())
                      }
                    />
                  </Form.Item>
                </Col>
              </Row>
              <Form.Item
                label="需要解决的业务问题"
                name="problem"
                rules={[
                  { required: true, message: "请描述需要解决的业务问题" },
                ]}
              >
                <Input.TextArea rows={3} />
              </Form.Item>
              <Row gutter={14}>
                <Col span={8}>
                  <Form.Item
                    label="数据对象"
                    name="objects"
                    rules={[{ required: true }]}
                  >
                    <Select
                      mode="multiple"
                      maxTagCount="responsive"
                      options={demandObjectOptions.map((value) => ({ value }))}
                    />
                  </Form.Item>
                </Col>
                <Col span={8}>
                  <Form.Item
                    label="覆盖地域"
                    name="regions"
                    rules={[{ required: true }]}
                  >
                    <Select
                      mode="multiple"
                      showSearch
                      maxTagCount="responsive"
                      options={domesticRegions.map((value) => ({ value }))}
                    />
                  </Form.Item>
                </Col>
                <Col span={8}>
                  <Form.Item
                    label="更新频率"
                    name="frequency"
                    rules={[{ required: true }]}
                  >
                    <Select
                      options={demandFrequencyOptions.map((value) => ({
                        value,
                      }))}
                    />
                  </Form.Item>
                </Col>
              </Row>
              <Row gutter={14}>
                <Col span={16}>
                  <Form.Item label="重点字段或指标" name="fields">
                    <Select
                      mode="tags"
                      tokenSeparators={[",", "，"]}
                      options={[
                        "统一社会信用代码",
                        "注册时间",
                        "所属行业",
                        "控股类型",
                        "企业资质",
                        "招投标记录",
                        "司法风险",
                        "舆情事件",
                        "关联关系",
                        "经营异常",
                      ].map((value) => ({ value }))}
                    />
                  </Form.Item>
                </Col>
                <Col span={8}>
                  <Form.Item label="期望使用时间" name="deadline">
                    <Input />
                  </Form.Item>
                </Col>
              </Row>
              <Button
                type="primary"
                size="large"
                icon={<SearchOutlined />}
                onClick={inspect}
              >
                匹配库内产品
              </Button>
              <span className="demand-form-tip">
                系统同时比较业务任务、数据对象、地域、频率、字段和交付要求
              </span>
            </Form>
          </Card>
        </Col>
        <Col span={9}>
          <Card
            title="库内产品推荐"
            extra={
              matched ? (
                <Tag color="green">已找到 {recommendations.length} 个</Tag>
              ) : (
                <Tag>等待匹配</Tag>
              )
            }
            className="demand-match-card"
          >
            {matched ? (
              <>
                <div className="match-lead">
                  <RobotOutlined />
                  <span>
                    <b>优先复用现有产品</b>
                    <small>
                      请先确认推荐产品是否能够解决需求，避免重复报送和重复建设。
                    </small>
                  </span>
                </div>
                <div className="demand-recommendations">
                  {recommendations.map((x, i) => (
                    <div className={i === 0 ? "best" : ""} key={x.id}>
                      <div className="rec-head">
                        <span>
                          <Tag color={i === 0 ? "red" : "default"}>{x.fit}</Tag>
                          <b>{x.name}</b>
                          <small>
                            {x.id} · {x.scope}
                          </small>
                        </span>
                        <Progress
                          type="circle"
                          percent={x.score}
                          size={48}
                          strokeColor="#c7000b"
                        />
                      </div>
                      <p>{x.reason}</p>
                      <Space>
                        <Button size="small" onClick={() => go("app-products")}>
                          查看详情
                        </Button>
                        <Button
                          size="small"
                          type={i === 0 ? "primary" : "default"}
                          onClick={() => useProduct(x.name)}
                        >
                          直接使用
                        </Button>
                      </Space>
                    </div>
                  ))}
                </div>
                <div className="gap-confirm">
                  <Checkbox
                    checked={confirmedGap}
                    onChange={(e) => setConfirmedGap(e.target.checked)}
                  >
                    上述产品均无法完整满足，仍需提交差异需求
                  </Checkbox>
                  {confirmedGap && (
                    <>
                      <Input.TextArea
                        rows={2}
                        placeholder="请说明现有产品不能满足的具体差异，如缺少字段、地域覆盖不足或更新时效不符"
                      />
                      <Button danger block onClick={() => setSubmitted(true)}>
                        上交未满足的数据需求
                      </Button>
                    </>
                  )}
                </div>
              </>
            ) : (
              <Empty
                image={Empty.PRESENTED_IMAGE_SIMPLE}
                description="填写或调整需求后，点击“匹配库内产品”"
              >
                <span className="empty-help">
                  匹配前不会直接进入采购或建设流程
                </span>
              </Empty>
            )}
          </Card>
        </Col>
      </Row>
    </div>
  );
}

function ProcurementManage({ go }: { go: (v: View) => void }) {
  const [active, setActive] = useState("demands");
  const [stage, setStage] = useState(4);
  const [open, setOpen] = useState(false);
  const [issueStatus, setIssueStatus] = useState("询价中");
  const [fullFlow, setFullFlow] = useState(false);
  const [query, setQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("全部状态");
  const flowNodes = [
    "需求确认",
    "存量查重",
    "采购对象确认",
    "供应商询价",
    "方案比选",
    "样本验证",
    "合规评估",
    "采购立项",
    "验收入库",
  ];
  const phaseCurrent =
    stage <= 1 ? 0 : stage <= 3 ? 1 : stage <= 5 ? 2 : stage <= 7 ? 3 : 4;
  const phaseItems = [
    { title: "需求核定", description: "确认缺口" },
    { title: "采购准备", description: "对象与询价" },
    { title: "方案验证", description: "比选与样本" },
    { title: "合规立项", description: "评估与审批" },
    { title: "验收入库", description: "形成资源" },
  ];
  const demands = [
    {
      id: "DR-2026-0803-016",
      name: "北京专精特新企业营销数据需求",
      org: "北京市分行",
      scene: "客户营销",
      product: "浙江分行专精特新企业筛选产品",
      coverage: 88,
      gap: "司法事件摘要等4项字段",
      priority: "高",
      status: "缺口已确认",
    },
    {
      id: "DR-2026-0801-011",
      name: "供应链风险关系数据需求",
      org: "江苏省分行",
      scene: "风险监测",
      product: "广东核心企业供应链风险监测",
      coverage: 62,
      gap: "二级供应商关系与交易强度",
      priority: "高",
      status: "样本验证",
    },
    {
      id: "DR-2026-0729-008",
      name: "区域产业政策事件数据需求",
      org: "四川省分行",
      scene: "产业研究",
      product: "产业事件驱动融资机会识别",
      coverage: 91,
      gap: "地方政策标签",
      priority: "中",
      status: "存量覆盖",
    },
  ];
  const issues = [
    {
      id: "PR-2026-0042",
      name: "企业司法风险摘要字段授权扩围",
      source: "DR-2026-0803-016",
      product: "浙江分行专精特新企业筛选产品",
      object: "司法事件摘要、裁判要旨等4项标准字段",
      supplier: "法信数据",
      method: "现有API合同扩围",
      budget: "24万元/年",
      owner: "王琳",
      deadline: "2026-08-12",
      status: issueStatus,
    },
    {
      id: "PR-2026-0039",
      name: "供应链二级关系数据试采",
      source: "DR-2026-0801-011",
      product: "广东核心企业供应链风险监测",
      object: "二级供应商关系、交易强度、关系置信度",
      supplier: "数联智库",
      method: "竞争性询价＋样本试用",
      budget: "48万元/年",
      owner: "陈嘉敏",
      deadline: "2026-08-18",
      status: "样本验证",
    },
    {
      id: "PR-2026-0035",
      name: "地方专精特新名单更新服务",
      source: "CP-2026-0731-01",
      product: "浙江分行专精特新企业筛选产品",
      object: "省级名单月度增量与资质有效期",
      supplier: "浙江省分行报送",
      method: "地方公共数据接入",
      budget: "0元",
      owner: "林晓舟",
      deadline: "2026-08-06",
      status: "待验收入库",
    },
  ];
  const matches = (x: any) =>
    `${x.id}${x.name}${x.product || ""}${x.supplier || ""}${x.org || ""}`
      .toLowerCase()
      .includes(query.trim().toLowerCase()) &&
    (statusFilter === "全部状态" || x.status === statusFilter);
  const shownDemands = demands.filter(matches);
  const shownIssues = issues.filter(matches);
  const advance = () => {
    const next = Math.min(stage + 1, 8);
    setStage(next);
    setIssueStatus(
      next >= 8
        ? "已验收入库"
        : next >= 6
          ? "合规评估"
          : next >= 5
            ? "样本验证"
            : "询价中",
    );
    message.success(
      next >= 8
        ? "资源已验收入库，可供产品建设调用"
        : "采购事项已推进至下一节点",
    );
  };
  const demandColumns: any[] = [
    {
      title: "业务需求",
      dataIndex: "name",
      width: "27%",
      render: (x: string, r: any) => (
        <div className="procurement-primary">
          <a
            onClick={() => {
              setActive("issues");
              setOpen(true);
            }}
          >
            {x}
          </a>
          <small>
            {r.id} · {r.org} · {r.scene}
          </small>
        </div>
      ),
    },
    {
      title: "存量复用判断",
      width: "27%",
      render: (_: any, r: any) => (
        <div className="procurement-stack">
          <span>{r.product}</span>
          <Progress
            percent={r.coverage}
            size="small"
            strokeColor={r.coverage > 85 ? "#15945b" : "#d77c00"}
          />
        </div>
      ),
    },
    {
      title: "待补缺口",
      dataIndex: "gap",
      width: "19%",
      render: (x: string, r: any) => (
        <div className="procurement-stack">
          <span>{x}</span>
          <small>
            优先级：
            <Tag color={r.priority === "高" ? "red" : "orange"}>
              {r.priority}
            </Tag>
          </small>
        </div>
      ),
    },
    {
      title: "当前状态",
      dataIndex: "status",
      width: "13%",
      render: (x: string) => <StatusTag status={x} />,
    },
    {
      title: "下一步",
      width: "14%",
      render: (_: any, r: any) => (
        <Button
          size="small"
          onClick={() =>
            r.status === "存量覆盖"
              ? message.info("该需求由存量资源满足，无需采购")
              : setActive("issues")
          }
        >
          {r.status === "存量覆盖" ? "复用方案" : "采购事项"}
        </Button>
      ),
    },
  ];
  const issueColumns: any[] = [
    {
      title: "采购事项",
      dataIndex: "name",
      width: "26%",
      render: (x: string, r: any) => (
        <div className="procurement-primary">
          <a onClick={() => setOpen(true)}>{x}</a>
          <small>
            {r.id} · 来源 {r.source}
          </small>
        </div>
      ),
    },
    {
      title: "采购对象与供应商",
      width: "28%",
      render: (_: any, r: any) => (
        <div className="procurement-stack">
          <span>{r.object}</span>
          <small>候选供应商：{r.supplier}</small>
        </div>
      ),
    },
    {
      title: "采购方案",
      width: "18%",
      render: (_: any, r: any) => (
        <div className="procurement-stack">
          <span>{r.method}</span>
          <small>预算：{r.budget}</small>
        </div>
      ),
    },
    {
      title: "办理信息",
      width: "13%",
      render: (_: any, r: any) => (
        <div className="procurement-stack">
          <span>{r.owner}</span>
          <small>{r.deadline}</small>
        </div>
      ),
    },
    {
      title: "状态 / 操作",
      width: "15%",
      render: (_: any, r: any) => (
        <Space direction="vertical" size={5}>
          <StatusTag status={r.status} />
          <Button size="small" type="primary" onClick={() => setOpen(true)}>
            办理
          </Button>
        </Space>
      ),
    },
  ];
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            采购需求管理 <DemoTag />
          </h1>
          <p>
            从业务缺口确认采购对象，贯通询价、样本验证、合规评估、立项和验收入库。
          </p>
        </div>
        <Space>
          <Button onClick={() => message.info("已导出采购需求台账")}>
            导出台账
          </Button>
          <Button
            type="primary"
            onClick={() => {
              setActive("issues");
              setOpen(true);
            }}
          >
            新建采购事项
          </Button>
        </Space>
      </div>
      <Row gutter={[14, 14]} className="procurement-metrics">
        {[
          ["待分析业务需求", 18, "本周新增 6"],
          ["存量资源可覆盖", 11, "避免重复采购"],
          ["在办采购事项", 7, "2项临近时限"],
          ["预计节约金额", "186万", "较整包采购"],
        ].map((x) => (
          <Col span={6} key={x[0]}>
            <MetricCard
              title={String(x[0])}
              value={x[1]}
              trend={String(x[2])}
            />
          </Col>
        ))}
      </Row>
      <div className="procurement-board">
        <Card className="procurement-nav">
          <b>采购管理工作区</b>
          {[
            ["demands", "需求归集与查重", "18项"],
            ["issues", "采购事项台账", "7项"],
            ["progress", "在办流程跟踪", "5项"],
            ["acceptance", "验收与资源入库", "3项"],
          ].map((x) => (
            <button
              className={active === x[0] ? "active" : ""}
              onClick={() => setActive(x[0])}
              key={x[0]}
            >
              <span>{x[1]}</span>
              <em>{x[2]}</em>
            </button>
          ))}
          <div className="procurement-guide">
            <SafetyCertificateOutlined />
            <b>采购边界</b>
            <p>
              只有经存量查重确认无法覆盖的数据、授权或服务缺口，才转为采购事项。
            </p>
          </div>
        </Card>
        <Card
          className="procurement-main"
          title={
            active === "demands"
              ? "需求归集与存量查重"
              : active === "issues"
                ? "采购事项台账"
                : active === "progress"
                  ? "在办流程跟踪"
                  : "验收与资源入库"
          }
          extra={
            <Space wrap>
              <Input
                allowClear
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                prefix={<SearchOutlined />}
                placeholder="搜索需求、产品或供应商"
              />
              <Select
                value={statusFilter}
                onChange={setStatusFilter}
                popupMatchSelectWidth={180}
                options={[
                  "全部状态",
                  "存量覆盖",
                  "缺口已确认",
                  "询价中",
                  "样本验证",
                  "合规评估",
                  "待验收入库",
                  "已验收入库",
                ].map((value) => ({ value }))}
              />
            </Space>
          }
        >
          {active === "demands" && (
            <>
              <div className="procurement-tip">
                <RobotOutlined />
                <span>
                  系统已按字段、主体、地域、时间、更新频率和使用授权完成库内查重，并合并3项同类需求。
                </span>
                <Tag color="green">本批次复用率 81%</Tag>
              </div>
              <Table
                className="procurement-table"
                tableLayout="fixed"
                pagination={false}
                rowKey="id"
                dataSource={shownDemands}
                columns={demandColumns}
              />
            </>
          )}
          {active === "issues" && (
            <>
              <div className="procurement-tip">
                <FileDoneOutlined />
                <span>
                  每项采购事项必须明确来源需求、关联产品、采购对象、预算、责任人和交付标准。
                </span>
                <Tag color="red">临期 2 项</Tag>
              </div>
              <Table
                className="procurement-table"
                tableLayout="fixed"
                pagination={false}
                rowKey="id"
                dataSource={shownIssues}
                columns={issueColumns}
              />
            </>
          )}
          {active === "progress" && (
            <div className="procurement-track">
              <div className="track-summary">
                <div>
                  <b>PR-2026-0042 · 企业司法风险摘要字段授权扩围</b>
                  <span>已完成 {Math.min(stage, 8)}/8 个流转节点</span>
                </div>
                <Progress
                  percent={Math.round((stage / 8) * 100)}
                  showInfo={false}
                  strokeColor="#c7000b"
                />
                <Button size="small" onClick={() => setFullFlow(!fullFlow)}>
                  {fullFlow ? "收起完整流程" : "查看完整流程"}
                </Button>
              </div>
              <Steps
                className="phase-steps"
                size="small"
                responsive
                current={phaseCurrent}
                items={phaseItems}
              />
              {fullFlow && (
                <div className="full-flow">
                  <Timeline
                    items={flowNodes.map((title, i) => ({
                      color:
                        i < stage ? "green" : i === stage ? "blue" : "gray",
                      children: (
                        <span>
                          <b>{title}</b>
                          {i === stage && <Tag color="blue">当前节点</Tag>}
                        </span>
                      ),
                    }))}
                  />
                </div>
              )}
              <Row gutter={[14, 14]}>
                <Col xs={24} xl={16}>
                  <Card
                    size="small"
                    title={
                      <Space>
                        <span>当前办理</span>
                        <StatusTag status={issueStatus} />
                      </Space>
                    }
                    extra={
                      <span className="track-deadline">
                        计划完成：2026-08-12
                      </span>
                    }
                  >
                    <div className="current-task">
                      <div className="current-task-index">{stage + 1}</div>
                      <div>
                        <small>当前节点</small>
                        <h3>{flowNodes[Math.min(stage, 8)]}</h3>
                        <p>
                          等待经办人提交供应商样本、字段字典、报价单与授权范围说明。
                        </p>
                      </div>
                    </div>
                    <Descriptions
                      bordered
                      size="small"
                      column={2}
                      items={[
                        {
                          label: "关联产品",
                          children: "浙江分行专精特新企业筛选产品",
                        },
                        { label: "责任人", children: "王琳（数据采购岗）" },
                        { label: "已停留", children: "1个工作日" },
                        {
                          label: "剩余时限",
                          children: <Tag color="orange">3个工作日</Tag>,
                        },
                      ]}
                    />
                    <Timeline
                      className="process-history"
                      items={[
                        {
                          color: "green",
                          children: "08-03 完成需求合并与存量资源查重",
                        },
                        {
                          color: "green",
                          children: "08-04 确认4项字段授权缺口",
                        },
                        {
                          color: "blue",
                          children: `当前：${issueStatus}，等待经办人提交节点材料`,
                        },
                      ]}
                    />
                    <Button type="primary" onClick={advance}>
                      {stage >= 8 ? "已完成验收" : "提交本节点并继续"}
                    </Button>
                  </Card>
                </Col>
                <Col xs={24} xl={8}>
                  <Card size="small" title="办理门禁与协同">
                    <List
                      size="small"
                      dataSource={[
                        "预算余额校验通过",
                        "供应商准入状态正常",
                        "数据安全分级待确认",
                        "产品建设可先使用脱敏样本",
                      ]}
                      renderItem={(x, i) => (
                        <List.Item>
                          <CheckCircleFilled
                            className={i === 2 ? "red" : "green"}
                          />
                          {x}
                        </List.Item>
                      )}
                    />
                    <Button block onClick={() => go("build")}>
                      返回关联产品建设
                    </Button>
                  </Card>
                </Col>
              </Row>
            </div>
          )}
          {active === "acceptance" && (
            <Table
              scroll={{ x: "max-content" }}
              pagination={false}
              rowKey="id"
              dataSource={shownIssues.filter(
                (x) => x.status === "待验收入库" || x.status === "已验收入库",
              )}
              columns={[
                { title: "待验收资源", dataIndex: "name" },
                { title: "关联产品", dataIndex: "product" },
                { title: "交付单位", dataIndex: "supplier" },
                { title: "交付内容", dataIndex: "object" },
                {
                  title: "验收检查",
                  render: () => (
                    <Space wrap>
                      <Tag color="green">字段完整</Tag>
                      <Tag color="green">样本通过</Tag>
                      <Tag color="orange">授权待确认</Tag>
                    </Space>
                  ),
                },
                {
                  title: "操作",
                  render: () => (
                    <Button
                      type="primary"
                      size="small"
                      onClick={() => go("resource-onboarding")}
                    >
                      验收并转资源接入
                    </Button>
                  ),
                },
              ]}
            />
          )}
        </Card>
      </div>
      <Drawer
        width={760}
        open={open}
        onClose={() => setOpen(false)}
        title="采购事项办理 · PR-2026-0042"
        extra={<StatusTag status={issueStatus} />}
      >
        <div className="procurement-detail-head">
          <div>
            <Tag color="red">字段授权扩围</Tag>
            <h2>企业司法风险摘要字段授权扩围</h2>
            <p>
              由产品建设中心资源查重自动生成，解决司法风险核验环节的摘要字段授权缺口。
            </p>
          </div>
          <Progress
            type="circle"
            percent={Math.round((stage / 8) * 100)}
            size={86}
            strokeColor="#c7000b"
          />
        </div>
        <Descriptions
          bordered
          column={2}
          items={[
            { label: "来源需求", children: "DR-2026-0803-016" },
            {
              label: "关联候选产品",
              children: (
                <a onClick={() => go("build")}>浙江分行专精特新企业筛选产品</a>
              ),
            },
            { label: "采购对象", children: "4项司法风险摘要标准字段", span: 2 },
            { label: "候选供应商", children: "法信数据" },
            { label: "预算上限", children: "24万元/年" },
            { label: "期望交付", children: "标准API＋字段字典＋授权说明" },
            { label: "计划完成", children: "2026-08-12" },
          ]}
        />
        <h3>本节点材料</h3>
        <Upload.Dragger>
          <p className="ant-upload-drag-icon">
            <CloudUploadOutlined />
          </p>
          <p>上传报价单、样本数据、接口文档或授权说明</p>
        </Upload.Dragger>
        <div className="drawer-actions">
          <Button onClick={() => message.success("事项草稿已保存")}>
            保存草稿
          </Button>
          <Button type="primary" onClick={advance}>
            提交本节点
          </Button>
        </div>
      </Drawer>
    </div>
  );
}

const opportunityRows = [
  {
    id: "OP-260806-01",
    name: "先进制造重大订单与设备更新融资机会",
    source: "政策事件＋重复自助组合",
    value: 94,
    common: 91,
    data: 88,
    cost: 62,
    priority: "P0",
    status: "待立项",
    owner: "公司金融部",
  },
  {
    id: "OP-260805-07",
    name: "供应链二级风险传导预警",
    source: "8家分行相似报送",
    value: 90,
    common: 86,
    data: 64,
    cost: 71,
    priority: "P0",
    status: "补充数据",
    owner: "风险管理部",
  },
  {
    id: "OP-260804-12",
    name: "园区重点企业扩产线索识别",
    source: "未满足搜索 126次",
    value: 83,
    common: 78,
    data: 92,
    cost: 45,
    priority: "P1",
    status: "待评审",
    owner: "机构金融部",
  },
  {
    id: "OP-260731-09",
    name: "科创企业知识产权质押机会",
    source: "业务人员反馈",
    value: 79,
    common: 72,
    data: 81,
    cost: 58,
    priority: "P1",
    status: "方案预研",
    owner: "普惠金融部",
  },
];

function OpportunityPool({ go }: { go: (v: View) => void }) {
  const [selected, setSelected] = useState(opportunityRows[0]);
  const [status, setStatus] = useState(selected.status);
  const score = Math.round(
    selected.value * 0.32 +
      selected.common * 0.22 +
      selected.data * 0.18 +
      (100 - selected.cost) * 0.14 +
      86 * 0.14,
  );
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            产品机会池 <DemoTag />
          </h1>
          <p>
            汇聚分行报送、业务反馈、高频找数、重复组合、政策事件与运行异常，形成可解释的建设优先级。
          </p>
        </div>
        <Space>
          <Button onClick={() => message.success("已刷新机会识别日志")}>
            刷新识别
          </Button>
          <Button
            type="primary"
            icon={<RobotOutlined />}
            onClick={() => message.success("系统新增识别3项潜在产品机会")}
          >
            智能发现机会
          </Button>
        </Space>
      </div>
      <Row gutter={[12, 12]} className="factory-metrics">
        {[
          ["待评估机会", 36, "本周新增 8"],
          ["跨区域共性", 17, "覆盖3家以上机构"],
          ["建议立项", 6, "综合评分≥85"],
          ["重复建设规避", "286万", "预计年度节约"],
        ].map((x) => (
          <Col xs={12} xl={6} key={String(x[0])}>
            <MetricCard
              title={String(x[0])}
              value={x[1]}
              trend={String(x[2])}
            />
          </Col>
        ))}
      </Row>
      <div className="factory-split">
        <Card title="机会优先级队列" extra={<Tag color="red">七维评分</Tag>}>
          <Table
            rowKey="id"
            pagination={false}
            dataSource={opportunityRows}
            onRow={(r) => ({
              onClick: () => {
                setSelected(r);
                setStatus(r.status);
              },
            })}
            columns={[
              {
                title: "产品机会",
                dataIndex: "name",
                render: (x: string, r: any) => (
                  <div className="cell-main">
                    <b>{x}</b>
                    <small>
                      {r.id} · {r.source}
                    </small>
                  </div>
                ),
              },
              { title: "价值", dataIndex: "value" },
              { title: "共性", dataIndex: "common" },
              { title: "数据", dataIndex: "data" },
              { title: "成本", dataIndex: "cost" },
              {
                title: "优先级",
                dataIndex: "priority",
                render: (x: string) => (
                  <Tag color={x === "P0" ? "red" : "orange"}>{x}</Tag>
                ),
              },
              {
                title: "状态",
                dataIndex: "status",
                render: (x: string) => <StatusTag status={x} />,
              },
            ]}
          />
        </Card>
        <Card className="opportunity-detail" title="机会评估卡">
          <Tag color="red">
            {selected.priority} · 综合 {score}分
          </Tag>
          <h2>{selected.name}</h2>
          <p>
            {selected.source}触发，拟由{selected.owner}牵头验证。
          </p>
          <div className="radar-list">
            {[
              ["业务价值", selected.value],
              ["跨区域共性", selected.common],
              ["数据可得性", selected.data],
              ["流程嵌入度", 86],
              ["可验证性", 89],
            ].map((x) => (
              <div key={String(x[0])}>
                <span>{x[0]}</span>
                <Progress
                  percent={Number(x[1])}
                  size="small"
                  strokeColor="#c7000b"
                />
              </div>
            ))}
          </div>
          <Descriptions
            size="small"
            column={1}
            items={[
              { label: "机会来源", children: selected.source },
              {
                label: "建议交付",
                children: "名单＋融资需求判断＋客户经理任务",
              },
              { label: "建设路径", children: "母版适配＋能力组合＋地方参数" },
            ]}
          />
          <Select
            value={status}
            onChange={setStatus}
            options={["待评审", "待立项", "补充数据", "方案预研", "已立项"].map(
              (value) => ({ value }),
            )}
          />
          <Button
            type="primary"
            block
            onClick={() => {
              setStatus("已立项");
              message.success("已生成产品定义画布草案");
              go("build");
            }}
          >
            立项并进入产品定义
          </Button>
        </Card>
      </div>
    </div>
  );
}

function FactoryDashboard({ go }: { go: (v: View) => void }) {
  const stages = [
    { name: "机会发现", count: 36, time: "1.2天" },
    { name: "定义评审", count: 12, time: "1.8天" },
    { name: "组合建设", count: 9, time: "6.4天" },
    { name: "仿真验证", count: 7, time: "3.1天" },
    { name: "试点灰度", count: 5, time: "18天" },
    { name: "正式运营", count: 89, time: "持续" },
  ];
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            外部数据产品创新工厂 <DemoTag />
          </h1>
          <p>
            以统一资产底座为基础，贯通产品发现、定义、生产、验证、发布、运营和复制。
          </p>
        </div>
        <Space>
          <Button onClick={() => message.success("驾驶舱数据已更新")}>
            更新数据
          </Button>
          <Button type="primary" onClick={() => go("opportunities")}>
            查看机会池
          </Button>
        </Space>
      </div>
      <Row gutter={[12, 12]} className="factory-metrics">
        {[
          ["平均建设周期", "13.6天", "同比缩短 38%"],
          ["存量资产复用率", "78%", "较上季 +9pct"],
          ["试点通过率", "84%", "5项正在灰度"],
          ["跨区域复制", "24项", "覆盖31家机构"],
        ].map((x) => (
          <Col xs={12} xl={6} key={String(x[0])}>
            <MetricCard
              title={String(x[0])}
              value={x[1]}
              trend={String(x[2])}
            />
          </Col>
        ))}
      </Row>
      <Card className="factory-flow" title="产品生产线">
        <div>
          {stages.map((x, i) => (
            <React.Fragment key={x.name}>
              <button
                onClick={() =>
                  go(
                    i === 0
                      ? "opportunities"
                      : i === 4
                        ? "pilot"
                        : i === 5
                          ? "operations"
                          : "build",
                  )
                }
              >
                <span>0{i + 1}</span>
                <b>{x.name}</b>
                <em>{x.count}项</em>
                <small>平均 {x.time}</small>
              </button>
              {i < stages.length - 1 && <i>→</i>}
            </React.Fragment>
          ))}
        </div>
      </Card>
      <Row gutter={[12, 12]}>
        <Col xs={24} xl={15}>
          <Card title="建设效能趋势">
            <ReactECharts
              style={{ height: 300 }}
              option={{
                tooltip: { trigger: "axis" },
                legend: { data: ["建设周期（天）", "复用率（%）"], top: 4, left: "center", itemGap: 28 },
                grid: { left: 48, right: 48, top: 54, bottom: 36, containLabel: true },
                xAxis: {
                  type: "category",
                  data: ["3月", "4月", "5月", "6月", "7月", "8月"],
                },
                yAxis: [{ type: "value" }, { type: "value", max: 100 }],
                series: [
                  {
                    name: "建设周期（天）",
                    type: "bar",
                    data: [28, 25, 22, 19, 16, 13.6],
                    itemStyle: { color: "#c7000b" },
                  },
                  {
                    name: "复用率（%）",
                    type: "line",
                    yAxisIndex: 1,
                    data: [52, 57, 61, 68, 73, 78],
                    itemStyle: { color: "#d89614" },
                  },
                ],
              }}
            />
          </Card>
        </Col>
        <Col xs={24} xl={9}>
          <Card title="关键门禁">
            <List
              dataSource={[
                "2项机会待业务价值确认",
                "1项模型漂移超过阈值",
                "3项灰度产品待扩大范围",
                "4项复制包等待地方适配",
              ]}
              renderItem={(x, i) => (
                <List.Item
                  actions={[
                    <Button
                      key="x"
                      size="small"
                      onClick={() =>
                        go(i === 2 ? "pilot" : i === 1 ? "quality" : "build")
                      }
                    >
                      处理
                    </Button>,
                  ]}
                >
                  <Badge status={i === 1 ? "error" : "warning"} />
                  {x}
                </List.Item>
              )}
            />
          </Card>
        </Col>
      </Row>
    </div>
  );
}

function PilotManagement({ go }: { go: (v: View) => void }) {
  const [progress, setProgress] = useState(62);
  const [decision, setDecision] = useState("继续灰度");
  const pilots = [
    {
      id: "PL-2608-03",
      name: "先进制造重大订单融资机会",
      org: "北京、江苏、四川分行",
      period: "2026-07-20—08-20",
      group: "试验 1,200 / 对照 1,200",
      hit: "68%",
      convert: "12.4%",
      status: "灰度运行",
    },
    {
      id: "PL-2607-11",
      name: "供应链风险传导预警",
      org: "广东、浙江分行",
      period: "2026-07-01—08-15",
      group: "试验 860 / 对照 840",
      hit: "74%",
      convert: "风险核验 31%",
      status: "待复盘",
    },
  ];
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            试点与灰度管理 <DemoTag />
          </h1>
          <p>
            测试通过后进入业务试点，以对照样本、任务承载和发布门禁验证产品是否真正有效。
          </p>
        </div>
        <Button
          type="primary"
          onClick={() => message.success("已创建新的灰度计划")}
        >
          新建试点计划
        </Button>
      </div>
      <Row gutter={[12, 12]} className="factory-metrics">
        {[
          ["在试产品", 5, "覆盖9家分行"],
          ["试验样本", "4,860", "对照样本4,720"],
          ["平均命中率", "71%", "较基线+18pct"],
          ["待发布决策", 3, "本周完成"],
        ].map((x) => (
          <Col xs={12} xl={6} key={String(x[0])}>
            <MetricCard
              title={String(x[0])}
              value={x[1]}
              trend={String(x[2])}
            />
          </Col>
        ))}
      </Row>
      <Card title="试点台账">
        <Table
          rowKey="id"
          pagination={false}
          dataSource={pilots}
          columns={[
            {
              title: "试点产品",
              dataIndex: "name",
              render: (x: string, r: any) => (
                <div className="cell-main">
                  <b>{x}</b>
                  <small>{r.id}</small>
                </div>
              ),
            },
            { title: "试点机构", dataIndex: "org" },
            { title: "观察周期", dataIndex: "period" },
            { title: "样本设计", dataIndex: "group" },
            { title: "命中率", dataIndex: "hit" },
            { title: "业务转化", dataIndex: "convert" },
            {
              title: "状态",
              dataIndex: "status",
              render: (x: string) => <StatusTag status={x} />,
            },
          ]}
        />
      </Card>
      <Row gutter={[12, 12]}>
        <Col xs={24} xl={15}>
          <Card
            title="PL-2608-03 · 观察进度"
            extra={<Tag color="blue">灰度运行</Tag>}
          >
            <Progress percent={progress} strokeColor="#c7000b" />
            <div className="pilot-gates">
              {[
                ["数据质量", 96, "通过"],
                ["名单命中", 68, "通过"],
                ["业务转化", 82, "通过"],
                ["任务承载", 76, "通过"],
                ["接口稳定", 99, "通过"],
                ["责任确认", 60, "待确认"],
              ].map((x) => (
                <div key={String(x[0])}>
                  <span>{x[0]}</span>
                  <b>{x[1]}%</b>
                  <StatusTag status={String(x[2])} />
                </div>
              ))}
            </div>
            <Space wrap>
              <Button onClick={() => setProgress(Math.min(100, progress + 10))}>
                记录一轮观察
              </Button>
              <Button onClick={() => message.success("已生成试点评估报告")}>
                生成评估报告
              </Button>
              <Button
                danger
                onClick={() => message.warning("已保留当前版本，可执行回滚")}
              >
                回滚演练
              </Button>
            </Space>
          </Card>
        </Col>
        <Col xs={24} xl={9}>
          <Card title="发布决策">
            <Radio.Group
              value={decision}
              onChange={(e) => setDecision(e.target.value)}
            >
              <Space direction="vertical">
                <Radio value="扩大试点">扩大试点</Radio>
                <Radio value="继续灰度">继续灰度</Radio>
                <Radio value="整改后重测">整改后重测</Radio>
                <Radio value="终止">终止</Radio>
              </Space>
            </Radio.Group>
            <Input.TextArea
              rows={4}
              defaultValue="业务指标达到阶段门槛，责任部门确认后扩大至华东地区。"
            />
            <Button
              type="primary"
              block
              onClick={() => {
                message.success(`决策已提交：${decision}`);
                go("build-publish");
              }}
            >
              提交发布门禁审批
            </Button>
          </Card>
        </Col>
      </Row>
    </div>
  );
}

function QualityCenter({ go }: { go: (v: View) => void }) {
  const checks = [
    {
      stage: "事前准入",
      item: "授权范围与更新频率",
      score: 98,
      status: "通过",
      owner: "资源管理岗",
    },
    {
      stage: "事中生产",
      item: "主体关联准确率",
      score: 96.8,
      status: "通过",
      owner: "能力运营岗",
    },
    {
      stage: "事中生产",
      item: "字段映射异常率",
      score: 92,
      status: "关注",
      owner: "产品建设岗",
    },
    {
      stage: "事后巡检",
      item: "模型特征漂移 PSI",
      score: 71,
      status: "异常",
      owner: "模型管理岗",
    },
    {
      stage: "事后巡检",
      item: "结果抽检与误报率",
      score: 88,
      status: "关注",
      owner: "业务运营岗",
    },
  ];
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            质量控制中心 <DemoTag />
          </h1>
          <p>
            将事前预防、事中控制和事后巡检嵌入产品生产线，统一管理数据、规则、模型、结果与调用质量。
          </p>
        </div>
        <Button
          type="primary"
          onClick={() => message.success("全量质量巡检任务已启动")}
        >
          启动巡检
        </Button>
      </div>
      <Row gutter={[12, 12]} className="factory-metrics">
        {[
          ["质量规则", 128, "启用117项"],
          ["今日异常", 7, "高风险1项"],
          ["主体关联准确率", "96.8%", "抽样2.4万条"],
          ["平均修复时长", "4.2小时", "同比-31%"],
        ].map((x) => (
          <Col xs={12} xl={6} key={String(x[0])}>
            <MetricCard
              title={String(x[0])}
              value={x[1]}
              trend={String(x[2])}
            />
          </Col>
        ))}
      </Row>
      <div className="quality-chain">
        {["事前准入", "事中生产", "事后巡检"].map((x, i) => (
          <div key={x}>
            <span>0{i + 1}</span>
            <b>{x}</b>
            <p>
              {i === 0
                ? "授权、覆盖、时效、样本"
                : i === 1
                  ? "映射、异常、关联、冲突、稳定性"
                  : "抽检、漂移、误报、异常调用"}
            </p>
          </div>
        ))}
      </div>
      <Card title="质量门禁与处置">
        <Table
          rowKey={(r) => r.stage + r.item}
          pagination={false}
          dataSource={checks}
          columns={[
            { title: "控制阶段", dataIndex: "stage" },
            { title: "检查项", dataIndex: "item" },
            {
              title: "得分",
              dataIndex: "score",
              render: (x: number) => (
                <Progress
                  percent={x}
                  size="small"
                  strokeColor={x < 80 ? "#d4380d" : "#c7000b"}
                />
              ),
            },
            {
              title: "状态",
              dataIndex: "status",
              render: (x: string) => <StatusTag status={x} />,
            },
            { title: "责任人", dataIndex: "owner" },
            {
              title: "处置",
              render: (_: any, r: any) => (
                <Button
                  size="small"
                  onClick={() =>
                    message.success(
                      r.status === "异常"
                        ? "已创建模型重训与策略复核任务"
                        : "已进入检查证据",
                    )
                  }
                >
                  {r.status === "异常" ? "发起整改" : "查看证据"}
                </Button>
              ),
            },
          ]}
        />
      </Card>
      <div className="identity-note">
        <SafetyCertificateOutlined />
        <span>
          <b>质量可信度同步至产品详情</b>
          业务用户可查看数据更新时间、关键字段覆盖率、结果置信度、适用边界、已知缺陷和最近验证结果。
        </span>
        <Button onClick={() => go("product-detail")}>查看产品侧展示</Button>
      </div>
    </div>
  );
}

function TaskLoop() {
  const [rows, setRows] = useState([
    {
      id: "TK-260806-1842",
      company: "北京智造装备有限公司",
      trigger: "重大订单 2.6亿元",
      role: "客户经理",
      deadline: "今日17:00",
      status: "待确认",
      result: "—",
    },
    {
      id: "TK-260806-1798",
      company: "苏州精工自动化股份",
      trigger: "设备更新项目",
      role: "客户经理",
      deadline: "明日12:00",
      status: "已触达",
      result: "融资需求3,000万元",
    },
    {
      id: "TK-260805-1621",
      company: "成都链创科技有限公司",
      trigger: "扩产技改",
      role: "风险经理",
      deadline: "已完成",
      status: "已回流",
      result: "转入授信调查",
    },
  ]);
  const advance = (id: string) =>
    setRows(
      rows.map((r) =>
        r.id === id
          ? {
              ...r,
              status: r.status === "待确认" ? "已触达" : "已回流",
              result: r.status === "待确认" ? "待记录融资需求" : "转入授信调查",
            }
          : r,
      ),
    );
  return (
    <div>
      <div className="page-heading">
        <div>
          <h1>
            业务任务与结果回流 <DemoTag />
          </h1>
          <p>
            外数产品以业务动作结束：触发任务、分派岗位、记录核验和触达、进入授信或处置流程，并将结果回流产品评价。
          </p>
        </div>
        <Button
          type="primary"
          onClick={() => message.success("已按管户关系分派18条新任务")}
        >
          分派今日任务
        </Button>
      </div>
      <div className="task-loop-flow">
        {[
          "产业事件触发",
          "企业与风险核验",
          "生成融资机会",
          "客户经理确认",
          "触达与需求记录",
          "授信/处置结果回流",
        ].map((x, i) => (
          <React.Fragment key={x}>
            <span>
              <em>{i + 1}</em>
              {x}
            </span>
            {i < 5 && <b>→</b>}
          </React.Fragment>
        ))}
      </div>
      <Row gutter={[12, 12]} className="factory-metrics">
        {[
          ["今日生成", 186, "有效任务"],
          ["按时确认率", "91%", "目标≥90%"],
          ["融资需求", 38, "金额4.2亿元"],
          ["转入授信", 17, "转化率9.1%"],
        ].map((x) => (
          <Col xs={12} xl={6} key={String(x[0])}>
            <MetricCard
              title={String(x[0])}
              value={x[1]}
              trend={String(x[2])}
            />
          </Col>
        ))}
      </Row>
      <Card title="任务执行台账">
        <Table
          rowKey="id"
          pagination={false}
          dataSource={rows}
          columns={[
            {
              title: "企业与任务",
              dataIndex: "company",
              render: (x: string, r: any) => (
                <div className="cell-main">
                  <b>{x}</b>
                  <small>{r.id}</small>
                </div>
              ),
            },
            { title: "触发依据", dataIndex: "trigger" },
            { title: "接收岗位", dataIndex: "role" },
            { title: "完成时限", dataIndex: "deadline" },
            {
              title: "当前状态",
              dataIndex: "status",
              render: (x: string) => <StatusTag status={x} />,
            },
            { title: "业务结果", dataIndex: "result" },
            {
              title: "操作",
              render: (_: any, r: any) => (
                <Button
                  size="small"
                  disabled={r.status === "已回流"}
                  onClick={() => advance(r.id)}
                >
                  {r.status === "待确认" ? "确认并触达" : "回流结果"}
                </Button>
              ),
            },
          ]}
        />
      </Card>
      <Card title="回流数据如何驱动迭代" className="loop-attribution">
        <div>
          {[
            ["数据延迟", "调整数据源"],
            ["主体错配", "修正映射"],
            ["误报偏高", "优化规则/模型"],
            ["未及时处理", "调整任务路由"],
            ["转化偏低", "复核产品价值"],
            ["验证有效", "回沉能力与复制包"],
          ].map((x) => (
            <span key={x[0]}>
              <b>{x[0]}</b>
              <em>→</em>
              {x[1]}
            </span>
          ))}
        </div>
      </Card>
    </div>
  );
}

function Placeholder({ view, go }: { view: string; go: (v: View) => void }) {
  const map: any = {
    scenes: [
      "场景找数",
      "从客户拓展、授信调查、风险监测等业务场景定位产品与资源",
    ],
    relations: ["关联视图", "查看资源、产品、能力、规则和场景之间的关系"],
    follows: ["我的关注", "管理已关注的产品、资源和业务主题"],
    process: ["流程中心", "查看全部流程实例、审批记录与个人待办"],
    settings: ["系统配置", "维护组织、用户、流程模板和基础参数"],
    validation: ["验证发布", "集中管理待验证产品、测试报告与发布审批"],
  };
  const x = map[view] || ["功能页面", "该页面为演示范围内的辅助功能"];
  return (
    <Card>
      <Result
        icon={<AppstoreOutlined className="placeholder-icon" />}
        title={x[0]}
        subTitle={x[1]}
        extra={
          <Space>
            <Button
              onClick={() =>
                go(view.startsWith("app") ? "app-home" : "manage-home")
              }
            >
              返回首页
            </Button>
            <Button
              type="primary"
              onClick={() =>
                go(view === "validation" ? "build" : "app-products")
              }
            >
              进入核心演示流程
            </Button>
          </Space>
        }
      />
    </Card>
  );
}

const portalProcesses: Record<Exclude<Portal, "login">, Array<{ title: string; desc: string; view: View }>> = {
  manage: [
    { title: "资源准备", desc: "接入并核对数据", view: "resource-onboarding" },
    { title: "需求受理", desc: "明确任务与边界", view: "candidates" },
    { title: "能力建设", desc: "定义模块输入输出", view: "build-capability" },
    { title: "模块组装", desc: "按业务顺序编排", view: "build-assembly" },
    { title: "测试发布", desc: "验证后审批上架", view: "build-testing" },
    { title: "成品管理", desc: "查看版本与范围", view: "finished" },
  ],
  branch: [
    { title: "明确任务", desc: "说明场景与对象", view: "branch-build" },
    { title: "选择资源", desc: "调用资源和能力", view: "branch-build" },
    { title: "试运行", desc: "检查过程与结果", view: "branch-build" },
    { title: "在线报送", desc: "提交完整建设档案", view: "report" },
    { title: "查看进度", desc: "跟踪受理与发布", view: "branch-submissions" },
  ],
  app: [
    { title: "明确任务", desc: "选择业务场景", view: "scenes" },
    { title: "选择产品", desc: "匹配适用产品", view: "app-products" },
    { title: "查看依据", desc: "了解数据与边界", view: "product-detail" },
    { title: "设置参数", desc: "确定地区和条件", view: "workbench" },
    { title: "结果与行动", desc: "研判并生成任务", view: "follows" },
  ],
};

function PortalProcessBar({ portal, view, go }: { portal: Exclude<Portal, "login">; view: View; go: (v: View) => void }) {
  const steps = portalProcesses[portal];
  const aliases: Record<string, string> = {
    build: "build-capability",
    "build-agent": "build-capability",
    "build-publish": "build-testing",
    "manage-home": "resource-onboarding",
    "app-home": "scenes",
    "app-resources": "scenes",
    suppliers: "scenes",
    feedback: "follows",
    "branch-home": "branch-build",
    "branch-feedback": "branch-submissions",
  };
  const activeView = aliases[view] || view;
  const activeIndex = Math.max(0, steps.findIndex((step) => step.view === activeView));
  return (
    <nav className={`portal-process portal-process-${portal}`} aria-label="业务流程">
      <div className="portal-process-title"><small>当前主流程</small><b>{portal === "app" ? "从业务任务到行动" : portal === "branch" ? "从地方建设到报送" : "从资源和能力到标准产品"}</b></div>
      <div className="portal-process-steps">
        {steps.map((step, index) => (
          <button key={`${step.title}-${index}`} className={index === activeIndex ? "active" : index < activeIndex ? "done" : ""} onClick={() => go(step.view)}>
            <span>{index < activeIndex ? <CheckCircleFilled /> : index + 1}</span>
            <b>{step.title}</b>
            <small>{step.desc}</small>
          </button>
        ))}
      </div>
    </nav>
  );
}

function Shell() {
  const [reportDemo, setReportDemo] = useState(false);
  const [portal, setPortal] = useState<Portal>("login");
  const [role, setRole] = useState<RoleKey>("business");
  const [view, setView] = useState<View>("manage-home");
  const [collapsed, setCollapsed] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [published, setPublished] = useState(false);
  const [deposited, setDeposited] = useState(false);
  const [newReview, setNewReview] = useState(false);
  const [productName, setProductName] = useState("全国产业链机会与风险识别");
  useEffect(() => {
    setSubmitted(localStorage.getItem("icbc-submitted") === "true");
    setPublished(localStorage.getItem("icbc-published") === "true");
    setDeposited(localStorage.getItem("icbc-deposited") === "true");
    setNewReview(localStorage.getItem("icbc-review") === "true");
  }, []);
  useEffect(
    () =>
      setProductName(
        localStorage.getItem("icbc-product-name") || "全国产业链机会与风险识别",
      ),
    [],
  );
  const renameProduct = (name: string) => {
    setProductName(name);
    localStorage.setItem("icbc-product-name", name);
  };
  const portalHome = (p: Portal) =>
    p === "app" ? "app-home" : p === "branch" ? "branch-home" : "manage-home";
  const login = (r: RoleKey) => {
    setRole(r);
    setPortal(roleMap[r].portal);
    setView(portalHome(roleMap[r].portal));
  };
  const switchRole = (r: RoleKey) => {
    setRole(r);
    setPortal(roleMap[r].portal);
    setView(portalHome(roleMap[r].portal));
    message.success(`已切换为${roleMap[r].name}`);
  };
  const switchPortal = (p: Portal) => {
    setPortal(p);
    setView(portalHome(p));
  };
  const usePublishedProduct = (name: string) => {
    localStorage.setItem("icbc-workbench-product", name);
    setPortal("app");
    setView("workbench");
    message.success(`已进入“${name}”运行工作台`);
  };
  if (reportDemo) return <ReportDemo onExit={() => setReportDemo(false)} />;
  if (portal === "login") return <Login onLogin={login} onDemo={() => setReportDemo(true)} />;
  const go = (v: View) => setView(v);
  const render = () => {
    if (portal === "branch") {
      if (view === "branch-home")
        return (
          <BranchHome go={go} submitted={submitted} published={published} />
        );
      if (view === "branch-build") return <BranchProductStudio go={go} />;
      if (view === "report")
        return (
          <ReportForm
            go={go}
            onSubmitted={() => {
              setSubmitted(true);
            }}
          />
        );
      if (view === "branch-submissions")
        return <BranchSubmissions published={published} go={go} />;
      if (view === "branch-feedback")
        return <BranchFeedback published={published} />;
      return <BranchHome go={go} submitted={submitted} published={published} />;
    }
    if (portal === "manage") {
      if (view === "manage-home") return <ManageHome go={go} />;
      if (view === "factory-dashboard") return <FactoryDashboard go={go} />;
      if (view === "opportunities") return <OpportunityPool go={go} />;
      if (view === "manage-resources") return <ResourceManagement />;
      if (view === "procurement") return <ProcurementManage go={go} />;
      if (view === "resource-onboarding") return <ResourceOnboarding go={go} />;
      if (view === "candidates")
        return (
          <CandidateZone
            submitted={submitted}
            productName={productName}
            onRename={renameProduct}
            goBuild={() => setView("build")}
          />
        );
      if (view === "build-agent") return <ProductBuildAgent go={go} />;
      if (
        view === "build" ||
        view === "build-capability" ||
        view === "build-assembly" ||
        view === "build-testing" ||
        view === "build-publish"
      )
        return (
          <BuildCenter
            key={view}
            initialTab={
              view === "build-capability"
                ? "match"
                : view === "build-assembly"
                  ? "canvas"
                  : view === "build-testing"
                    ? "testing"
                : view === "build-publish"
                  ? "publish"
                  : "agent"
            }
            productName={productName}
            published={published}
            onPublish={() => setPublished(true)}
            onReset={() => {
              setPublished(false);
              setDeposited(false);
            }}
            go={go}
            onUse={usePublishedProduct}
          />
        );
      if (view === "pilot") return <PilotManagement go={go} />;
      if (view === "quality") return <QualityCenter go={go} />;
      if (view === "assets") return <Assets deposited={deposited} go={go} />;
      if (view === "finished")
        return (
          <FinishedProducts
            published={published}
            productName={productName}
            onUse={usePublishedProduct}
          />
        );
      if (view === "operations") return <Operations newReview={newReview} />;
      if (view === "task-loop") return <TaskLoop />;
      if (view === "process") return <ProcessCenter go={go} />;
      if (view === "settings") return <SystemConfig />;
      return <Placeholder view={view} go={go} />;
    }
    if (view === "app-home") return <AppHome go={go} />;
    if (view === "app-resources") return <ResourceManagement appMode />;
    if (view === "data-demand") return <DataDemand go={setView} />;
    if (view === "suppliers") return <SupplierCenter />;
    if (view === "app-products")
      return <AppProducts published={published} go={go} />;
    if (view === "product-detail")
      return <ProductDetail published={published} go={go} />;
    if (view === "scenes") return <SceneFinder go={go} />;
    if (view === "follows") return <FollowCenter go={go} />;
    if (view === "workbench") return <Workbench go={go} />;
    if (view === "feedback")
      return <Feedback onReview={() => setNewReview(true)} />;
    return <Placeholder view={view} go={go} />;
  };
  const menu =
    portal === "manage"
      ? manageMenu
      : portal === "branch"
        ? branchMenu
        : appMenu;
  const portalName =
    portal === "manage"
      ? "总行建设管理门户"
      : portal === "branch"
        ? "分行建设与报送门户"
        : "业务应用门户";
  return (
    <Layout className="platform">
      <TopHeader
        role={role}
        onRole={switchRole}
        onLogout={() => setPortal("login")}
        portal={portal}
        onPortal={switchPortal}
      />
      <Layout>
        <Sider
          width={216}
          collapsedWidth={70}
          collapsed={collapsed}
          className="side-nav"
        >
          <div className={`portal-label${collapsed ? " collapsed" : ""}`}>
            <span aria-hidden={collapsed}>{portalName}</span>
            <Tooltip
              title={collapsed ? `展开${portalName}` : "收起侧边栏"}
              placement="right"
            >
              <button
                aria-label={collapsed ? `展开${portalName}` : "收起侧边栏"}
                onClick={() => setCollapsed(!collapsed)}
              >
                {collapsed ? <MenuUnfoldOutlined /> : <MenuFoldOutlined />}
              </button>
            </Tooltip>
          </div>
          <Menu
            mode="inline"
            selectedKeys={[view]}
            items={menu}
            onClick={({ key }) => setView(key)}
          />
          <div className={`side-bottom${collapsed ? " collapsed" : ""}`}>
            <Tag color="red">
              {portal === "manage"
                ? "总行管理端"
                : portal === "branch"
                  ? "分行建设端"
                  : "业务应用端"}
            </Tag>
            <DemoTag />
          </div>
        </Sider>
        <Content className="main-content">
          <PortalProcessBar portal={portal} view={view} go={go} />
          {render()}
          <div className="page-foot">
            全部内容均为演示数据 · 中国工商银行外部数据产品管理平台概念原型
          </div>
        </Content>
      </Layout>
    </Layout>
  );
}

export default function Home() {
  return (
    <ConfigProvider
      theme={{
        token: {
          colorPrimary: "#c7000b",
          borderRadius: 4,
          colorInfo: "#c7000b",
          fontFamily: '"Microsoft YaHei","PingFang SC",Arial,sans-serif',
        },
        components: {
          Button: { controlHeight: 34 },
          Card: { headerHeight: 46 },
          Menu: { itemBorderRadius: 2 },
        },
      }}
    >
      <App>
        <Shell />
      </App>
    </ConfigProvider>
  );
}
