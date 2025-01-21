
"""Created on the 05.Dec.2024
@author: Dhan Lord B. Fortela
@author contact: dhanlord.fortela@louisiana.edu

API for accessing AspenPlus RadFrac module to perform SAC RL
If you change it, update it, fix something just email me such that I can also update my version to keep it as coherent as possible

Note that that functions currently contained in this module are those for the RadFrac column module in AspenPlus.

Refrence:
These codes are based on the prior work of: Richard ten Hagen; email: Richardxtenxhagen@gmail.com; GitHub: https://github.com/YouMayCallMeJesus/AspenPlus-Python-Interface. This prior work demonstrated example codes for AspenPlus-Python Interacing using the Python package win32com.

"""

# Import required packages:
from fileinput import filename
import os
from re import A
from tokenize import String
from typing import Union, Dict, Literal
import win32com.client as win32
import numpy as np
import time



# Set of funcitons for the AspenPlus simulation called withon Python:

class Simulation():
    """Class which starts a Simulation interface instance
    
    Args:
        AspenFileName: Name of the Aspenfile on which you are working with
        WorkingDirectoryPath: Path to the Folder where we will be working
        VISIBITLITY: Toggles the opening and interactive running of the Aspen simulation
    """
    #AspenSimulation = win32.gencache.EnsureDispatch("Apwn.Document") # this seems like the old syntax
    AspenSimulation = win32.Dispatch("Apwn.Document") # this initializes the connection of Python with Windows and ApsenPlus application

    def __init__(self, AspenFileName:str, WorkingDirectoryPath:str, VISIBILITY:bool = True):
        print("The current Directory is :  ")
        print(os.getcwd())                      #Returns the Directory where it is currently working
        os.chdir(WorkingDirectoryPath)          #Changes the Directory to  ..../AspenSimulation
        print("The new Directory where you should also have your Aspen file is : ")
        print(os.getcwd())          
        self.AspenSimulation.InitFromArchive2(os.path.abspath(AspenFileName))
        print("The Aspen is active now. If you dont want to see aspen open again take VISIBITLY as False \n")
        self.AspenSimulation.Visible = VISIBILITY

    def Run(self) -> bool:
        """Runs simulation, if there is a problem it will rerun twice, returns boolean about successful convergence"""
        tries = 0
        converged = 0
        #iterations = 10
        #self.BLK.Elements("B1").Elements("Input").Elements("MAXOL").Value = iterations

        while tries != 2:
            start = time.time()
            self.AspenSimulation.Engine.Run2()
            print(f"Runtime = {time.time() - start}")
            # print(time.time() - start)
            converged = self.AspenSimulation.Tree.Elements("Data").Elements("Results Summary").Elements(
                           "Run-Status").Elements("Output").Elements("PER_ERROR").Value
            if converged == 0:
                converged = True
                break
            elif converged == 1:
                tries += 1
                converged = False
        return converged
    
    
    
    
    def CloseAspen(self):
        AspenFileName = self.Give_AspenDocumentName()
        print(AspenFileName)
        self.AspenSimulation.Close(os.path.abspath(AspenFileName))
        print("\nAspen should be closed now")

    #This just shortens the path you need to call for Streams and Blocks:
    @property
    def BLK(self):
        """Property: Defines Path to the Block node in Aspen File system. 
            
        Aspendocument is defined in the Class Simulation initialization
        """
        return self.AspenSimulation.Tree.Elements("Data").Elements("Blocks") 

    @property
    def STRM(self):
        """Property: Defines Path to the Streamnode node in Aspen File system. 
            
        Aspendocument is defined in the Class Simulation initialization
        """
        return self.AspenSimulation.Tree.Elements("Data").Elements("Streams")



    #Type definition to simplify the type hinting:
    Phnum = Literal[1,2,3]
    Ph = Literal["L", "V", "S"]





    def VisibilityChange(self,VISIBILITY: bool) -> None:
        """ De/Activates Aspensheet graphics from being rendered. 
        
        Args:
            Visibility: String "FALSE" for more speed or "TRUE" for manual usage of Aspen
        """
        self.AspenSimulation.Visible = VISIBILITY
    
    def SheetCheckIfInputsAreComplete(self) -> bool:
        """Check if all Inputs are given on the entire Sheet, returns "0x00002081 = HAP_RESULTS_SUCCESS|HAP_INPUT_COMPLETE|HAP_ENABLED"
        
        Checks if the Aspen Expert system thinks all necessary Inputs are given and the Simulation can be run
        
        Args:
            Blockname: String which contains the Name of the Block in Aspen
            return: TRUE or FALSE????????????
        """
        return self.AspenSimulation.COMPSTATUS

    def BlockCheckIfInputsAreComplete(self, Blockname: str) -> bool:
        """
        Checks if the Aspen Expert system thinks all necessary Inputs are given and the Simulation can be run
        
        Args:
            Blockname: String which contains the Name of the Block in Aspen
            return: TRUE or FALSE????????????
        """
        return self.BLK.Elements(Blockname).COMPSTATUS

    def StreamCheckIfInputsAreComplete(self, Streamname:str) -> bool:
        """
        Checks if the Aspen Expert system thinks all necessary Inputs are given and the Simulation can be run
        
        Args:
            Streamname: String which contains the Name of the Block in Aspen
            return: TRUE or FALSE????????????
        """
        return self.STRM.Elements(Streamname).COMPSTATUS


    def Give_AspenDocumentName(self) -> String:
        """Returns name of Aspen document"""
        return self.AspenSimulation.FullName

    def DialogSuppression(self, TrueOrFalse: bool) -> None:
        """Supresses Aspen Popups
        
        Args: 
            TrueOrFalse: can be True or False """
        self.AspenSimulation.SuppressDialogs = TrueOrFalse
        
    def EngineRun(self) -> None:
        """Runs Simulation, synonymous with pressing the playbutton"""
        self.AspenSimulation.Run2()
        
    def EngineStop(self) -> None:
        """Stops Simulation, synonymous to pressing the red square button"""
        self.AspenSimulation.Stop()
        
    def EngineReinit(self) -> None:
        """Reinitalizes the Entire Simulation, synonymous to pressing the Reset button

        Other possible functions you might need are: BlockReinit(Blockname), StreamReinit(Streamname)
        """
        self.AspenSimulation.Reinit()
    
    def Save(self) -> None:
        """Saves Current Simulation (.apw), Inputs and all Values connected to it."""
        self.AspenSimulation.Save()





########################################################################################################################################


###########         N               N       PPPPPPPPPPP         U               U       TTTTTTTTTTTTTTTTTTTTTTTT
    #               N  N            N       P           P       U               U                   T
    #               N    N          N       P           P       U               U                   T
    #               N      N        N       P           P       U               U                   T
    #               N       N       N       PPPPPPPPPPP         U               U                   T
    #               N         N     N       P                   U               U                   T
    #               N           N   N       P                   U               U                   T
    #               N             N N       P                   U               U                   T
###########         N               N       P                   UUUUUUUUUUUUUUUUU                   T


############################################################################################################################################
#    
#
#
###################################################################################################################
#####
#############------- Read values of Action Var values from the AspenPlus file. -------##########
#####
###################################################################################################################

    """
    These are the functions used to GET (or read) the values of the Action variables (column internal design parameters). We read these values tp be inputs to the next training step of the SAC RL.

    """
    
    # Get individual parameter states:
    # Assuming only TOP and BOT sections exist
    def BLK_RADFRAC_Get_TOP_DIAMETER(self, Blockname):
        return self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DIAM").Elements("INT-1").Elements("TOP").Value
    def BLK_RADFRAC_Get_TOP_TRAYSPACING(self, Blockname):
        return self.BLK.Elements(Blockname).Elements("Input").Elements("CA_TRAY_SPC").Elements("INT-1").Elements("TOP").Value
    def BLK_RADFRAC_Get_TOP_WEIRHT(self, Blockname):        
        return self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIR_HT").Elements("INT-1").Elements("TOP").Value
    def BLK_RADFRAC_Get_TOP_DC_CLEAR(self, Blockname):            
        return self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DC_CLEAR").Elements("INT-1").Elements("TOP").Value
    def BLK_RADFRAC_Get_TOP_WEIR_SIDE_LN(self, Blockname):            
        return self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIRLN_SD").Elements("INT-1").Elements("TOP").Value
    def BLK_RADFRAC_Get_TOP_WEIR_HT(self, Blockname):            
        return self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIR_HT").Elements("INT-1").Elements("TOP").Value
    
    # BOT
    def BLK_RADFRAC_Get_BOT_DIAMETER(self, Blockname):
        return self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DIAM").Elements("INT-1").Elements("BOT").Value
    def BLK_RADFRAC_Get_BOT_TRAYSPACING(self, Blockname):
        return self.BLK.Elements(Blockname).Elements("Input").Elements("CA_TRAY_SPC").Elements("INT-1").Elements("BOT").Value
    def BLK_RADFRAC_Get_BOT_WEIRHT(self, Blockname):        
        return self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIR_HT").Elements("INT-1").Elements("BOT").Value
    def BLK_RADFRAC_Get_BOT_DC_CLEAR(self, Blockname):            
        return self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DC_CLEAR").Elements("INT-1").Elements("BOT").Value
    def BLK_RADFRAC_Get_BOT_WEIR_SIDE_LN(self, Blockname):            
        return self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIRLN_SD").Elements("INT-1").Elements("BOT").Value
    def BLK_RADFRAC_Get_TOP_WEIR_HT(self, Blockname):            
        return self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIR_HT").Elements("INT-1").Elements("BOT").Value

#    
#
#
###################################################################################################################
#####
#############------- Set the values of Action Var values into the AspenPlus file -------##########
#####
###################################################################################################################

    """
    These are the functions used to SET the values of the Action variables (column internal design parameters). Note that the assigned values are by default the read values from the initial setting in the AspenPlues model file. These values are updated by the SAC RL during leanring usigng the following functions .

    """
    
    # Set RadFrac Column internals assuming we have TOP and BOT sections only
    # TOP section
    def BLK_RADFRAC_Set_TOP_DIAMETER(self, Blockname, ColDiam_Top):
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DIAM").Elements("INT-1").Elements("TOP").Value = ColDiam_Top
    def BLK_RADFRAC_Set_TOP_TRAYSPACING(self, Blockname, TraySpace_Top):
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_TRAY_SPC").Elements("INT-1").Elements("TOP").Value = TraySpace_Top
    def BLK_RADFRAC_Set_TOP_DC_CLEAR(self, Blockname, DowncomerClearance_Top):            
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DC_CLEAR").Elements("INT-1").Elements("TOP").Value = DowncomerClearance_Top
    def BLK_RADFRAC_Set_TOP_WEIR_SIDE_LN(self, Blockname, WeirLengthSide_Top):            
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIRLN_SD").Elements("INT-1").Elements("TOP").Value = WeirLengthSide_Top
    def BLK_RADFRAC_Set_TOP_WEIR_HT(self, Blockname, WeirHeight_Top):            
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIR_HT").Elements("INT-1").Elements("TOP").Value = WeirHeight_Top
    def BLK_RADFRAC_Set_TOP_HOLE_DIAM(self, Blockname, HoleDiam_Top):            
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_HOLE_DIAM").Elements("INT-1").Elements("TOP").Value = HoleDiam_Top
    
    # BOT section 
    def BLK_RADFRAC_Set_BOT_DIAMETER(self, Blockname, ColDiam_Bot):
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DIAM").Elements("INT-1").Elements("BOT").Value = ColDiam_Bot
    def BLK_RADFRAC_Set_BOT_TRAYSPACING(self, Blockname, TraySpace_Bot):
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_TRAY_SPC").Elements("INT-1").Elements("BOT").Value = TraySpace_Bot
    def BLK_RADFRAC_Set_BOT_DC_CLEAR(self, Blockname, DowncomerClearance_Bot):            
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DC_CLEAR").Elements("INT-1").Elements("BOT").Value = DowncomerClearance_Bot
    def BLK_RADFRAC_Set_BOT_WEIR_SIDE_LN(self, Blockname, WeirLengthSide_Bot):            
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIRLN_SD").Elements("INT-1").Elements("BOT").Value = WeirLengthSide_Bot
    def BLK_RADFRAC_Set_BOT_WEIR_HT(self, Blockname, WeirHeight_Bot):            
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIR_HT").Elements("INT-1").Elements("BOT").Value = WeirHeight_Bot
    def BLK_RADFRAC_Set_BOT_HOLE_DIAM(self, Blockname, HoleDiam_Bot):            
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_HOLE_DIAM").Elements("INT-1").Elements("BOT").Value = HoleDiam_Bot



#
#
###################################################################################################################
#####
#############------- Get ALL values of Action Var values from the AspenPlus file -------##########
#####
###################################################################################################################

    """
    This function is not used in the work but is provided here just in case future works require the use of a functon that reads all Action var at the same time.
    """

    def BLK_RADFRAC_GET_ME_ALL_INPUTS_BACK(self, Blockname:str)-> Dict[str, Union[str,float,int]]:
        """Retrieves all the Inputs and returns Dictionary with Values 
        
        Does not include all aspects of a Aspen Simulationsheet, for this look at Exports
        
        Args:
            Blockname: String which gives the name of Block.         
        """
        ColDiam_Top = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DIAM").Elements("INT-1").Elements("TOP").Value
        TraySpace_Top = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_TRAY_SPC").Elements("INT-1").Elements("TOP").Value
        WeirHeight_Top = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIR_HT").Elements("INT-1").Elements("TOP").Value
        DowncomerClearance_Top = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DC_CLEAR").Elements("INT-1").Elements("TOP").Value
        WeirLengthSide_Top = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIRLN_SD").Elements("INT-1").Elements("TOP").Value
        WeirLengthSide_Top = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIRLN_SD").Elements("INT-1").Elements("TOP").Value
        WeirHeight_Top = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIR_HT").Elements("INT-1").Elements("TOP").Value
        
                #PAGE 6         Column Internal Design - Bot        
        ColDiam_Bot = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DIAM").Elements("INT-1").Elements("BOT").Value
        TraySpace_Bot = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_TRAY_SPC").Elements("INT-1").Elements("BOT").Value
        WeirHeight_Bot = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIR_HT").Elements("INT-1").Elements("BOT").Value
        DowncomerClearance_Bot = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DC_CLEAR").Elements("INT-1").Elements("BOT").Value
        WeirLengthSide_Bot = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIRLN_SD").Elements("INT-1").Elements("BOT").Value
        WeirHeight_Bot = self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIR_HT").Elements("INT-1").Elements("BOT").Value    

        
        Dictionary = {
            #PAGE 5         Column Internal Design - Top
            "ColDiam_Top":ColDiam_Top,
            "TraySpace_Top":TraySpace_Top,
            "WeirHeight_Top":WeirHeight_Top,
            "DowncomerClearance_Top":DowncomerClearance_Top,
            "WeirLengthSide_Top":WeirLengthSide_Top,
            #PAGE 6         Column Internal Design - Bot        
            "ColDiam_Bot":ColDiam_Bot,
            "TraySpace_Bot":TraySpace_Bot,
            "WeirHeight_Bot":WeirHeight_Bot,
            "DowncomerClearance_Bot":DowncomerClearance_Bot,
            "WeirLengthSide_Bot":WeirLengthSide_Bot
            
        }
        return Dictionary

#
#
###################################################################################################################
#####
#############------- Set ALL values of Action Var values into the AspenPlus file -------##########
#####
###################################################################################################################

    """
    This function is not used in the work but is provided here just in case future works require the use of a functon that sets all Action var at the same time.

    """
    
    def BLK_RADFRAC_SET_ALL_INPUTS(self, Blockname:str, Dictionary: Dict[str, Union[str,float,int]]) -> None:
        """Takes Dictionary with Values set the ones which are given in Aspen. 
        
        
        Args:
            Blockname: String which gives the name of Block.  
            Dictionary: Dictionary which contains all the Input variables.       
        """
        
                #PAGE 5         Column Internal Design - Top
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DIAM").Elements("INT-1").Elements("TOP").Value = Dictionary.get("ColDiam_Top")
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_TRAY_SPC").Elements("INT-1").Elements("TOP").Value = Dictionary.get("TraySpace_Top")
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIR_HT").Elements("INT-1").Elements("TOP").Value = Dictionary.get("WeirHeight_Top")
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DC_CLEAR").Elements("INT-1").Elements("TOP").Value = Dictionary.get("DowncomerClearance_Top")
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIRLN_SD").Elements("INT-1").Elements("TOP").Value = Dictionary.get("WeirLengthSide_Top")
                #PAGE 6         Column Internal Design - Bot        
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DIAM").Elements("INT-1").Elements("BOT").Value = Dictionary.get("ColDiam_Bot")
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_TRAY_SPC").Elements("INT-1").Elements("BOT").Value = Dictionary.get("TraySpace_Bot")
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIR_HT").Elements("INT-1").Elements("BOT").Value = Dictionary.get("WeirHeight_Bot")
        self.BLK.Elements(Blockname).Elements("Input").Elements("CA_DC_CLEAR").Elements("INT-1").Elements("BOT").Value = Dictionary.get("DowncomerClearance_Bot")
        fself.BLK.Elements(Blockname).Elements("Input").Elements("CA_WEIRLN_SD").Elements("INT-1").Elements("BOT").Value = Dictionary.get("WeirLengthSide_Bot")





###########################################################################################################################################



#ooooooooooooo   UU           U      tttttttttttttttttt         PPPPPPPPPPP         U               U       TTTTTTTTTTTTTTTTTTTTTTTT
#o           o   UU           U              TT                 P           P       U               U                   T
#o           o   UU           U              TT                 P           P       U               U                   T
#o           o   UU           U              TT                 P           P       U               U                   T
#o           o   UU           U              TT                 PPPPPPPPPPP         U               U                   T
#o           o   UU           U              TT                 P                   U               U                   T
#ooooooooooooo   UUUUUUUUUUUUUU              TT                 P                   U               U                   T
                                                               #P                   UUUUUUUUUUUUUUUUU                   T


############################################################################################################################################


###################################################################################################################
#####
############# Return Target State Var Values: % apporach to flooding values from all stages in the column ##########
#####
###################################################################################################################
    """
    These are the main functions used in the paper:  retrieve % flooding from each stage and report Maximum value as State var value.
    """
    # Return only the target State Vaiables: Maximum % flooding at TOP and BOT sections
    # get Max % of flooding value at Bot section   
    def BLK_RADFRAC_Get_TOP_Max_Flooding(self, Blockname):
        
        
        TopFloodingApproachList = []
        TopStageFloodingLister = self.BLK.Elements(Blockname).Elements("Output").Elements("CA_FLD_FAC8").Elements("INT-1").Elements("TOP").Elements
        
        
        for stage in TopStageFloodingLister:
            StageName = stage.Name
            TopFloodingApproachList.append(self.BLK.Elements(Blockname).Elements("Output").Elements("CA_FLD_FAC8").Elements("INT-1").Elements("TOP").Elements(StageName).Value)

        return max(TopFloodingApproachList)

    # get Max % of  flooding value at Bot section    
    def BLK_RADFRAC_Get_BOT_Max_Flooding(self, Blockname):
        # BOT Flooding
        BotFloodingApproachList = []
        BotStageFloodingLister = self.BLK.Elements(Blockname).Elements("Output").Elements("CA_FLD_FAC8").Elements("INT-1").Elements("BOT").Elements

        for stage in BotStageFloodingLister:
            StageName = stage.Name
            BotFloodingApproachList.append(self.BLK.Elements(Blockname).Elements("Output").Elements("CA_FLD_FAC8").Elements("INT-1").Elements("BOT").Elements(StageName).Value)
            
        return max(BotFloodingApproachList)



###################################################################################################
#####
############# Return ALl: % apporach to flooding values from all stages in the column #############
#####
###################################################################################################

    """
    This is not really used in the paper but it is provided here for reference.
    """

    
    def BLK_RADFRAC_GET_OUTPUTS(self, Blockname:str) -> Dict[str, Union[str,float,int]]:
        """Retrieves all Output variables for given Block and returns Dictionary of Values
            
            Args:
                Blockname: String which gives the name of Block.         
        """
        
        # TOP Flooding
        TopFloodingApproachList = []
        
        TopStageFloodingLister = self.BLK.Elements(Blockname).Elements("Output").Elements("CA_FLD_FAC8").Elements("INT-1").Elements("TOP").Elements
        
        
        for stage in TopStageFloodingLister:
            StageName = stage.Name
            TopFloodingApproachList.append(self.BLK.Elements(Blockname).Elements("Output").Elements("CA_FLD_FAC8").Elements("INT-1").Elements("TOP").Elements(StageName).Value)


        # BOT Flooding
        BotFloodingApproachList = []
        
        BotStageFloodingLister = self.BLK.Elements(Blockname).Elements("Output").Elements("CA_FLD_FAC8").Elements("INT-1").Elements("BOT").Elements

        for stage in BotStageFloodingLister:
            StageName = stage.Name
            BotFloodingApproachList.append(self.BLK.Elements(Blockname).Elements("Output").Elements("CA_FLD_FAC8").Elements("INT-1").Elements("BOT").Elements(StageName).Value)
        
        
        Dictionary = {
            
            "TopFloodingApproachList":TopFloodingApproachList,
            "BotFloodingApproachList":BotFloodingApproachList
            
        }
        return Dictionary
