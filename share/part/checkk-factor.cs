namespace tools;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;

public class checkk_factor
{
    public class CheckResult
    {
        public int BendCount { get; set; }
        public int ErrorCount { get; set; }
        public int WarningCount { get; set; }
        public bool IsSheetMetal { get; set; }
    }

    private static bool TryGetPoint75ThicknessSuggestion(double thicknessMm, out double suggestedThicknessMm)
    {
        // 按 0.01mm 精度处理，避免浮点误差导致的误报。
        double normalized = Math.Round(thicknessMm, 2);
        double nearestInteger = Math.Round(normalized, 0);
        if (Math.Abs(normalized - nearestInteger) > 0.05)
        {
            suggestedThicknessMm = 0;
            return false;
        }

        // 按工艺约束：整数板厚提示改为对应 .75 规格（如 3 -> 2.75）。
        suggestedThicknessMm = Math.Round(nearestInteger - 0.25, 2);
        return suggestedThicknessMm > 0;
    }


    public static bool Process_CustomBendAllowance(string modelname, CustomBendAllowance swCustBend, double BendRadius,
        string FeatureName, double thickness, double angle)
    {
        var debuct_factor = Math.Round(swCustBend.BendDeduction * 1000.0 , 2);
        var allow_factor = Math.Round( swCustBend.BendAllowance * 1000.0 , 2);
        double bendradius_limit = 6;

        // 检查大半径折弯的 K 因子设置 (R>=2mm 时必须使用 Type=2 且 KFactor=0.5)
        if (BendRadius >= bendradius_limit)
        {
            bool isTypeInvalid = swCustBend.Type != 2;
            bool isKFactorInvalid = swCustBend.Type == 2 && swCustBend.KFactor != 0.5;
            
            if (isTypeInvalid || isKFactorInvalid)
            {
                string reason = isTypeInvalid ? "未使用 K 因子类型" : "K 因子不等于 0.5";
                Console.WriteLine($"k 因子错误，{modelname}+{FeatureName},原因:{reason},当前值：Type={swCustBend.Type}, K 因子={swCustBend.KFactor}");
                return true;
            }

            Debug.Print("      扣除倍数      = " + debuct_factor);
            Debug.Print("      补偿换扣除 = " + allow_factor);
            Console.WriteLine($"k因子正确,{modelname}+{FeatureName}");
            return false;
        }

        else if (Math.Abs(angle - 90) > 0.5 && Math.Abs(swCustBend.KFactor - 0.35) > 0.05)
        {
            Console.WriteLine($"非 90 度折弯，k 因子错误，{modelname}+{FeatureName},触发条件：角度≠90°且 K 因子≠0.35,当前值：角度={angle}°, K 因子={swCustBend.KFactor}");
            return true;
        }

        else if (swCustBend.Type == 4 && (debuct_factor < 1.65*thickness || debuct_factor > 1.75*thickness))
        {
            Console.WriteLine($"k 因子错误，{modelname}+{FeatureName},触发条件：Type=4 且扣除倍数超出范围 [{ 1.65*thickness}-{1.75*thickness}],当前值：{debuct_factor}");
            return true;
        }


        else if (swCustBend.Type == 3 && (allow_factor < 0.25*thickness+2*BendRadius|| allow_factor > 0.35*thickness+2*BendRadius))
        {
            Console.WriteLine($"k 因子错误，{modelname}+{FeatureName},触发条件：Type=3 且补偿换扣除超出范围 [{0.25*thickness+2*BendRadius}-{0.35*thickness+2*BendRadius}],当前值：{allow_factor}");
            return true;
        }
        else if (BendRadius < bendradius_limit && swCustBend.Type == 2 && Math.Abs(swCustBend.KFactor - 0.35) > 0.05)
        {
            Console.WriteLine($"k 因子错误，{modelname}+{FeatureName},触发条件：R<{BendRadius}mm 且 Type=2 且 K 因子超出范围 [0.3-0.4],当前值：{swCustBend.KFactor}");
            return true;
        }

        

     
        else
        {
            Debug.Print("      扣除倍数      = " + debuct_factor);
            Debug.Print("      补偿换扣除 = " + allow_factor);
        
            Console.WriteLine($"k因子正确,{modelname}+{FeatureName}");
            return false;
        }
    }

    public static void PrintBendDebug(CustomBendAllowance swCustBend, double thickness, double bendRadius, double angle)
    {
        Debug.Print("      BendAllowance    = " + swCustBend.BendAllowance * 1000.0 + " mm");
        Debug.Print("      BendDeduction    = " + swCustBend.BendDeduction * 1000.0 + " mm");
        Debug.Print("      BendTableFile    = " + swCustBend.BendTableFile);
        Debug.Print("      KFactor          = " + swCustBend.KFactor);
        Debug.Print("      Type             = " + swCustBend.Type);
        Debug.Print("      thickness            = " + thickness);
        Debug.Print("      Radius = " + bendRadius + " mm");
        Debug.Print("      angle = " + angle + "°");
    }
  
    public static bool Process_OneBend(SldWorks swApp, ModelDoc2 swModel, Feature swFeat,double  thickness)
    {
        Debug.Print("    +" + swFeat.Name + " [" + swFeat.GetTypeName() + "]");

        OneBendFeatureData swOneBend = default(OneBendFeatureData);
        CustomBendAllowance swCustBend = default(CustomBendAllowance);

        swOneBend = (OneBendFeatureData)swFeat.GetDefinition();
        swCustBend = swOneBend.GetCustomBendAllowance(); 
        var angle=Math.Round(swOneBend.BendAngle* 180.0 / Math.PI,2);
        


        bool hasError = Process_CustomBendAllowance(swModel.GetTitle(), swCustBend, swOneBend.BendRadius * 1000.0, swFeat.Name, thickness, angle);
        if (hasError)
        {
            PrintBendDebug(swCustBend, thickness, swOneBend.BendRadius * 1000.0, angle);
        }
        return hasError;

    }


    static public CheckResult RunWithStats(SldWorks swApp, ModelDoc2 swModel)
    {
        CheckResult result = new CheckResult();

        // 检查是否为零件文档
        if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART)
        {
            Console.WriteLine($"跳过非零件文档: {swModel.GetTitle()} (类型: {swModel.GetType()})");
            return result;
        }
        
        var partdoc = (PartDoc)swModel;
        string modelName = swModel.GetTitle();

        // 按 Body 遍历：先取实体 body，再在 body 内遍历特征（与 exportdwg2_body / benddim 一致）
        object[] bodyObjects = (object[])partdoc.GetBodies2((int)swBodyType_e.swSolidBody, false);
        if (bodyObjects == null || bodyObjects.Length == 0)
        {
            Debug.WriteLine($"零件 '{modelName}' 没有可检查的实体 body");
            return result;
        }

        bool hasSheetMetalBody = false;
        int bendCount = 0;
        int errorCount = 0;

        foreach (object bodyObject in bodyObjects)
        {
            Body2 body = (Body2)bodyObject;
            object[] features = (object[])body.GetFeatures();
            if (features == null || features.Length == 0)
            {
                continue;
            }

            double thickness = 0;
            bool isSheetMetalBody = false;
            foreach (object featureObject in features)
            {
                Feature feature = (Feature)featureObject;
                if (feature.GetTypeName2() == "SheetMetal")
                {
                    SheetMetalFeatureData swSheetMetalData = (SheetMetalFeatureData)feature.GetDefinition();
                    thickness = swSheetMetalData.Thickness * 1000;
                    isSheetMetalBody = true;
                    hasSheetMetalBody = true;
                    result.IsSheetMetal = true;
                    break;
                }
            }

            if (!isSheetMetalBody)
            {
                continue;
            }

            if (TryGetPoint75ThicknessSuggestion(thickness, out double suggestedThickness))
            {
                string reminder = $"板厚提醒：{modelName}+{body.Name} 当前板厚 {Math.Round(thickness, 2)}mm 为整数，建议改为对应 .75 板厚（建议 {suggestedThickness:F2}mm）。";
                Console.WriteLine(reminder);
                result.WarningCount++;
                try
                {
                    swApp.SendMsgToUser(reminder);
                }
                catch
                {
                    // 某些批量场景下消息框调用可能失败，忽略不影响主流程。
                }
            }

            Debug.WriteLine($"开始检查钣金 body: {modelName}+{body.Name}, 厚度: {thickness}mm");

            foreach (object featureObject in features)
            {
                Feature swFeature = (Feature)featureObject;
                Feature swSubFeat = (Feature)swFeature.GetFirstSubFeature();
                while (swSubFeat != null)
                {
                    // 只处理 OneBend 特征
                    if (swSubFeat.GetTypeName() == "OneBend")
                    {
                        bendCount++;
                        try
                        {
                            bool hasError = Process_OneBend(swApp, swModel, swSubFeat, thickness);
                            if (hasError)
                            {
                                errorCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            Console.WriteLine($"检查折弯特征 '{swSubFeat.Name}' 时出错: {ex.Message}");
                        }
                    }
                    
                    swSubFeat = (Feature)swSubFeat.GetNextSubFeature();
               
                }
            }
        }

        if (!hasSheetMetalBody)
        {
            Debug.WriteLine($"零件 '{modelName}' 不是钣金件，跳过检查");
            return result;
        }
        
        Debug.WriteLine($"零件 '{modelName}' 检查完成: 共 {bendCount} 个折弯特征, {errorCount} 个错误");
        result.BendCount = bendCount;
        result.ErrorCount = errorCount;
        return result;
    }

    static public int run(SldWorks swApp, ModelDoc2 swModel)
    {
        return RunWithStats(swApp, swModel).BendCount;
    }
}