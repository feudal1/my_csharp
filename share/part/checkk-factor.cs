namespace tools;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;

public class checkk_factor
{


    public static void Process_CustomBendAllowance(string modelname, CustomBendAllowance swCustBend, double BendRadius,
        string FeatureName, double thickness, double angle)
    {
        var debuct_factor = Math.Round(swCustBend.BendDeduction * 1000.0 / thickness, 2);
        var a2d = Math.Round((2 * thickness - swCustBend.BendAllowance * 1000.0) / thickness, 2);
  

        // 检查大半径折弯的 K 因子设置 (R>=2mm 时必须使用 Type=2 且 KFactor=0.5)
        if (BendRadius >= 3)
        {
            bool isTypeInvalid = swCustBend.Type != 2;
            bool isKFactorInvalid = swCustBend.Type == 2 && swCustBend.KFactor != 0.5;
            
            if (isTypeInvalid || isKFactorInvalid)
            {
                string reason = isTypeInvalid ? "未使用 K 因子类型" : "K 因子不等于 0.5";
                Console.WriteLine($"k 因子错误，{modelname}+{FeatureName},原因:{reason},当前值：Type={swCustBend.Type}, K 因子={swCustBend.KFactor}");
            }
        }



        else if (swCustBend.Type == 4 && (debuct_factor < 1.6 || debuct_factor > 1.82))
        {
            Console.WriteLine($"k 因子错误，{modelname}+{FeatureName},触发条件：Type=4 且扣除倍数超出范围 [1.6-1.82],当前值：{debuct_factor}");
        }


        else if (swCustBend.Type == 3 && (a2d < 1.6 || a2d > 1.82))
        {
            Console.WriteLine($"k 因子错误，{modelname}+{FeatureName},触发条件：Type=3 且补偿换扣除超出范围 [1.6-1.82],当前值：{a2d}");
        }
        else if (BendRadius < 3 && swCustBend.Type == 2 && (swCustBend.KFactor < 0.2 || swCustBend.KFactor > 0.3))
        {
            Console.WriteLine($"k 因子错误，{modelname}+{FeatureName},触发条件：R<{BendRadius}mm 且 Type=2 且 K 因子超出范围 [0.2-0.3],当前值：{swCustBend.KFactor}");
        }

        
    
        else if (Math.Abs(angle - 90) > 0.5 && Math.Abs(swCustBend.KFactor - 0.25) > 0.05)
        {
            Console.WriteLine($"非 90 度折弯，k 因子错误，{modelname}+{FeatureName},触发条件：角度≠90°且 K 因子≠0.25,当前值：角度={angle}°, K 因子={swCustBend.KFactor}");
        }
        else
        {
            Debug.Print("      扣除倍数      = " + debuct_factor);
            Debug.Print("      补偿换扣除 = " + a2d);
        
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
        var partdoc = (PartDoc)swModel;
        var bodys = (object[])partdoc.GetBodies2((int)swBodyType_e.swSolidBody, false);

        foreach (var objbody in bodys)
        {
            var body = (Body2)objbody;

            double thickness = 0;
            object[] features = (object[])body.GetFeatures();
            foreach (object objFeature in features)
            {
 
                Feature swFeature = (Feature)objFeature;
                switch (swFeature.GetTypeName())
                {
                    case "SheetMetal":
                        SheetMetalFeatureData swSheetMetalData = (SheetMetalFeatureData)swFeature.GetDefinition();
                        thickness = swSheetMetalData.Thickness * 1000;
                        break;
                }
                var swSubFeat = (Feature)swFeature.GetFirstSubFeature();
 
                while ((swSubFeat != null))
                {
                    
                    switch (swSubFeat.GetTypeName())
                    {
                      
                            
                         
                        case "OneBend":
                            Process_OneBend(swApp, swModel, swSubFeat, thickness);

                            break;

                    }

                    swSubFeat = (Feature)swSubFeat.GetNextSubFeature();
                }







            }
        }
        return 0;
    }
}