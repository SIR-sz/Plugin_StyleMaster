using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using CadAtlasManager.Core;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media;

// [必选] 注册命令类，使 AutoCAD 能够识别该程序集中的 CommandMethod
[assembly: CommandClass(typeof(StyleMaster.MainTool))]

namespace StyleMaster
{
    /// <summary>
    /// StyleMaster 插件入口类
    /// 负责与 CadAtlasManager 宿主对接及命令行启动逻辑
    /// </summary>
    public class MainTool : ICadTool
    {
        // 授权标记：用于判断是否通过主程序合法启动
        private static bool _isAuthorized = false;

        #region --- ICadTool 接口实现 ---

        /// <summary>
        /// 插件在面板上显示的名称
        /// </summary>
        public string ToolName => "彩平大师 StyleMaster";

        /// <summary>
        /// 插件图标 (Segoe MDL2 Assets 编码)
        /// </summary>
        public string IconCode => "\uE790"; // 调色板图标

        /// <summary>
        /// 插件功能描述
        /// </summary>
        public string Description => "基于图层规则的自动化彩平填充工具，支持 Hatch 与 Image 混合渲染。";

        /// <summary>
        /// 插件分类
        /// </summary>
        public string Category { get; set; } = "景观渲染";

        /// <summary>
        /// 插件预览图 (由主程序自动加载)
        /// </summary>
        public ImageSource ToolPreview { get; set; }

        /// <summary>
        /// 验证宿主程序 GUID
        /// </summary>
        public bool VerifyHost(Guid hostGuid)
        {
            return hostGuid == new Guid("A7F3E2B1-4D5E-4B8C-9F0A-1C2B3D4E5F6B");
        }

        /// <summary>
        /// 主程序面板点击后的执行入口
        /// </summary>
        public void Execute()
        {
            // 标记为已授权
            _isAuthorized = true;
            ShowUIInternal();
        }

        #endregion

        #region --- 命令行入口 ---

        /// <summary>
        /// 命令行启动入口 (命令名: MPC)
        /// </summary>
        [CommandMethod("StyleMaster")]
        [CommandMethod("MPC")]
        public void MainCommandEntry()
        {
#if STANDALONE || DEBUG
            // 调试模式和独立版不检查授权
            ShowUIInternal();
#else
            // Release 模式下必须从主程序启动
            if (_isAuthorized)
            {
                ShowUIInternal();
            }
            else
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.Editor.WriteMessage("\n[错误] StyleMaster 为授权版插件，请通过主程序面板启动。");
                }
            }
#endif
        }

        #endregion

        #region --- 内部私有逻辑 ---

        /// <summary>
        /// 统一启动逻辑：包含资源目录初始化与 UI 显示
        /// </summary>
        private void ShowUIInternal()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            // 1. 初始化资源目录逻辑
            InitializeResources();

            // 2. 打印欢迎信息
            doc.Editor.WriteMessage($"\n[{ToolName}] 正在初始化环境...");

            // 3. 启动 UI (调用 UI 层静态方法)
            UI.MainControlWindow.ShowTool();
        }

        /// <summary>
        /// 初始化资源目录：在插件目录下确保存在 .\Resources\Materials\
        /// </summary>
        private void InitializeResources()
        {
            try
            {
                // 获取当前 DLL 所在目录
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string rootDir = Path.GetDirectoryName(assemblyPath);

                // 组合目标路径
                string materialPath = Path.Combine(rootDir, "Resources", "Materials");

                // 如果不存在则创建
                if (!Directory.Exists(materialPath))
                {
                    Directory.CreateDirectory(materialPath);
                }
            }
            catch (System.Exception ex)
            {
                // 仅在命令行提示错误，不中断运行
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage($"\n[警告] 资源目录创建失败: {ex.Message}");
            }
        }

        #endregion
    }
}