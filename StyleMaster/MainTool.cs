using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;
using System.Reflection;

// [必选] 注册命令类，使 AutoCAD 能够识别该程序集中的 CommandMethod
[assembly: CommandClass(typeof(StyleMaster.MainTool))]

namespace StyleMaster
{
    /// <summary>
    /// StyleMaster 插件入口类
    /// 改为独立运行版本，不再依赖 CadAtlasManager.Core 接口
    /// </summary>
    public class MainTool
    {
        #region --- 命令行入口 ---

        /// <summary>
        /// 命令行启动入口 (命令名: StyleMaster 或 MPC)
        /// </summary>
        [CommandMethod("StyleMaster")]
        [CommandMethod("MPC")]
        public void MainCommandEntry()
        {
            // 独立运行版本直接显示 UI
            ShowUIInternal();
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
            doc.Editor.WriteMessage("\n[StyleMaster] 正在初始化环境并启动插件...");

            // 3. 启动 UI (调用 UI 层静态方法)
            UI.MainControlWindow.ShowTool();
        }

        /// <summary>
        /// 初始化插件资源目录：确保 Patterns、.hatch_thumbs 和 Materials 文件夹完整存在
        /// </summary>
        private void InitializeResources()
        {
            try
            {
                // 获取当前 DLL 所在目录
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                string rootDir = Path.GetDirectoryName(assemblyPath);

                // 定义所有需要创建的路径
                string patternsPath = Path.Combine(rootDir, "Resources", "Patterns");
                string thumbsPath = Path.Combine(patternsPath, ".hatch_thumbs");
                string materialPath = Path.Combine(rootDir, "Resources", "Materials");

                // 逐级检查并创建
                string[] paths = { patternsPath, thumbsPath, materialPath };

                foreach (var path in paths)
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                }
            }
            catch (System.Exception ex)
            {
                // 仅在命令行提示错误，不中断运行
                Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage($"\n[警告] 资源目录自动创建失败: {ex.Message}");
            }
        }

        #endregion
    }
}