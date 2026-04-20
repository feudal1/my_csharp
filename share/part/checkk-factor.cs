namespace tools;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;

public class checkk_factor
{


    public static void Process_CustomBendAllowance(string modelname, CustomBendAllowance swCustBend, double BendRadius,
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
            }
        }

        else if (Math.Abs(angle - 90) > 0.5 && Math.Abs(swCustBend.KFactor - 0.35) > 0.05)
        {
            Console.WriteLine($"非 90 度折弯，k 因子错误，{modelname}+{FeatureName},触发条件：角度≠90°且 K 因子≠0.35,当前值：角度={angle}°, K 因子={swCustBend.KFactor}");
        }

        else if (swCustBend.Type == 4 && (debuct_factor < 1.65*thickness || debuct_factor > 1.75*thickness))
        {
            Console.WriteLine($"k 因子错误，{modelname}+{FeatureName},触发条件：Type=4 且扣除倍数超出范围 [{ 1.65*thickness}-{1.75*thickness}],当前值：{debuct_factor}");
        }


        else if (swCustBend.Type == 3 && (allow_factor < 0.25*thickness+2*BendRadius|| allow_factor > 0.35*thickness+2*BendRadius))
        {
            Console.WriteLine($"k 因子错误，{modelname}+{FeatureName},触发条件：Type=3 且补偿换扣除超出范围 [{0.25*thickness+2*BendRadius}-{0.35*thickness+2*BendRadius}],当前值：{allow_factor}");
        }
        else if (BendRadius < bendradius_limit && swCustBend.Type == 2 && Math.Abs(swCustBend.KFactor - 0.35) > 0.05)
        {
            Console.WriteLine($"k 因子错误，{modelname}+{FeatureName},触发条件：R<{BendRadius}mm 且 Type=2 且 K 因子超出范围 [0.3-0.4],当前值：{swCustBend.KFactor}");
        }

        

     
        else
        {
            Debug.Print("      扣除倍数      = " + debuct_factor);
            Debug.Print("      补偿换扣除 = " + allow_factor);
        
            Console.WriteLine($"k因子正确,{modelname}+{FeatureName}");
            return;
        }
        Debug.Print("      BendAllowance    = " + swCustBend.BendAllowance * 1000.0 + " mm");
        Debug.Print("      BendDeduction    = " + swCustBend.BendDeduction * 1000.0 + " mm");
        Debug.Print("      BendTableFile    = " + swCustBend.BendTableFile);
        Debug.Print("      KFactor          = " + swCustBend.KFactor);
        Debug.Print("      Type             = " + swCustBend.Type);
        Debug.Print("      thickness            = " + thickness);
        Debug.Print("      Radius = " + BendRadius + " mm");
        Debug.Print("      angle = " + angle + "°");

    }
  
    public static void Process_OneBend(SldWorks swApp, ModelDoc2 swModel, Feature swFeat,double  thickness)
    {
        Debug.Print("    +" + swFeat.Name + " [" + swFeat.GetTypeName() + "]");

        OneBendFeatureData swOneBend = default(OneBendFeatureData);
        CustomBendAllowance swCustBend = default(CustomBendAllowance);

        swOneBend = (OneBendFeatureData)swFeat.GetDefinition();
        swCustBend = swOneBend.GetCustomBendAllowance(); 
        var angle=Math.Round(swOneBend.BendAngle* 180.0 / Math.PI,2);
        


        Process_CustomBendAllowance( swModel.GetTitle(),swCustBend, swOneBend.BendRadius*1000.0,swFeat.Name ,thickness,angle);

    }


    static public int run(SldWorks swApp, ModelDoc2 swModel)
    {
        // 检查是否为零件文档
        if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART)
        {
            Console.WriteLine($"跳过非零件文档: {swModel.GetTitle()} (类型: {swModel.GetType()})");
            return 0;
        }
        
        var partdoc = (PartDoc)swModel;
        string modelName = swModel.GetTitle();
        
        // 优化1: 先快速检查是否为钣金件，避免不必要的遍历
        bool isSheetMetal = false;
        double thickness = 0;
        
        Feature firstFeature = (Feature)partdoc.FirstFeature();
        while (firstFeature != null)
        {
            if (firstFeature.GetTypeName2() == "SheetMetal")
            {
                SheetMetalFeatureData swSheetMetalData = (SheetMetalFeatureData)firstFeature.GetDefinition();
                thickness = swSheetMetalData.Thickness * 1000;
                isSheetMetal = true;
                break;
            }
            firstFeature = (Feature)firstFeature.GetNextFeature();
        }
        
        // 如果不是钣金件，直接返回
        if (!isSheetMetal)
        {
            Debug.WriteLine($"零件 '{modelName}' 不是钣金件，跳过检查");
            return 0;
        }
        
        Debug.WriteLine($"开始检查钣金件: {modelName}, 厚度: {thickness}mm");
        
        // 优化2: 只遍历一次特征树，找到所有折弯特征
        int bendCount = 0;
        int errorCount = 0;
        
        Feature swFeature = (Feature)partdoc.FirstFeature();
        while (swFeature != null)
        {
           
                Feature swSubFeat = (Feature)swFeature.GetFirstSubFeature();
                
                while (swSubFeat != null)
                {
                    // 只处理 OneBend 特征
                    if (swSubFeat.GetTypeName() == "OneBend")
                    {
                        bendCount++;
                        try
                        {
                            Process_OneBend(swApp, swModel, swSubFeat, thickness);
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            Console.WriteLine($"检查折弯特征 '{swSubFeat.Name}' 时出错: {ex.Message}");
                        }
                    }
                    
                    swSubFeat = (Feature)swSubFeat.GetNextSubFeature();
               
            }
            
            swFeature = (Feature)swFeature.GetNextFeature();
        }
        
        Debug.WriteLine($"零件 '{modelName}' 检查完成: 共 {bendCount} 个折弯特征, {errorCount} 个错误");
        return bendCount;
    }
}