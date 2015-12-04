%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% This program generates Non-Linear Programming MATLAB code from an existing Unisim/Hysis model. 
%
% in the form of:
%
% 	min f(x)    subject to:
%		 	lb <= x <= ub , x in R^n 	# Input bounds  
%			h(x) = 0 , h in R^m		# Equality constraints
%			g(x)<= 0 , g in R^k	 	# Inequality constraints
%			x0 				# Initial values for the 
%
% so it can be used for optimizing the process model using MATLAB's Optimization Toolbox.
% 
% It attempts to do that by connecting to an open Hysys/Unisim document through COM
% and if successiful, it generates all the required MATLAB code linked with Unisim/Hysis.
%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Neccessery steps for the program to run successifully :  
%
% 1. Create a Unisim/Hysis spreadsheet with a name "Objective" which should contain the value of the objective function in cell A2 e.g.
% 		A		B			
%	1	-		-		
%	2	f(x)		-
%
% 2. Create a Unisim/Hysis spreadsheet with a name "Constraints" which should have the following structure: 
% 		A		B		-	
%	1	-		-		-
%	2	g_1(x)		h_1(x)
%	3	g_2(x)		h_2(x)
%	|	  |		  |
%	k	g_k(x)		  |
%	|	  -		  |
%	m	  -		h_m(x)
%
% 3. Create a Unisim/Hysis spreadsheet with a name "Inputs" which should have the following structure: 
% 		A		B		C		D		E
%	1	-		-		-		-		-
%	2	x_1		ub(1) 		lb(1)		x0(1)	   Units of x_1
%	3 	x_2		ub(2) 		lb(2)		x0(2)	   Units of x_2
%	|	 |		  |		  |		  |		|
%	n	x_n		ub(n) 		lb(n)		x0(n)	   Units of x_n 
%
% Now you are ready !
%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% A few general comments:
% 
% As you may have noticed the first row in all the spreadsheets isn't used by the programm so it can be used for naming the columns e.g.
% 		A		B		C		D		E
%	1     Inputs	   Upper bounds	    Lower bounds   Initial Values     Units
%	2	x_1		ub(1) 		lb(1)		x0(1)		C
%	|	 |		  |		  |		  |	       Watt
%
% The program scans the Unisim/Hysis spreadsheet columns starting from the second row till it encounters an empty cell. 
% So the safest thing would be to use additional columns for calculations or comments but if you want to use the rest of the column cells
% for calculations, make sure to leave an empty cell at the end before using the rest of the cells.
%
% Also make sure that you specify the correct unit for each input (that would the unit that corresponds to the values that you use to specify the bounds) since
% the program scales all the inputs to [0,1]
%
%
% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% This program is free software: you can redistribute it and/or modify
% it under the terms of the GNU General Public License V3 as published by
% the Free Software Foundation
%
% This program is distributed in the hope that it will be useful,
% but WITHOUT ANY WARRANTY; without even the implied warranty of
% MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
% GNU General Public License for more details.

% For a copy of the GNU General Public License see <http://www.gnu.org/licenses/>.
%
% Copyright (C) 2012  Vlad Minasides 
%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%
%--------------------------------Program----------------------------------------------

% A Unisim simulation must be open before running this script.

clc
clear all
close all

% Declare global variables
global h;
global hyCase;
global f;

% Set the type of the model Unisim or Hysys
isUnisim = 1

%% Create links using COM 
h = actxserver('UnisimDesign.Application');

    hyCase = h.Activedocument;
f = hyCase.Flowsheet;


%% Read the Inputs spreadsheet

% Count the inputs
i=0;
while f.Operations.Item('Inputs').Cell(strcat(['A' int2str(i+2)])).CellVariable.IsKnown
    i=i+1;  
end

numberOfInputs=i;

if numberOfInputs<1
    error('Warning, no inputs');
end
   
inputIndices= [ 2:1:numberOfInputs+1;
                2:1:numberOfInputs+1];
inputIndices2= [ 1:1:numberOfInputs;
                2:1:numberOfInputs+1];
inputIndices3=[ 2:1:numberOfInputs+1;
                1:1:numberOfInputs;
                2:1:numberOfInputs+1];
inputIndices4= [ 1:1:numberOfInputs;
                2:1:numberOfInputs+1;
                1:1:numberOfInputs];

% Read the upper and lower bounds plus the initial values
inputLB=zeros(numberOfInputs,1);
inputUB=zeros(numberOfInputs,1);
initialUs=zeros(numberOfInputs,1);

parameter.inputUnits=cell(numberOfInputs,1);

for i=1:1:numberOfInputs
    inputLB(i)=f.Operations.Item('Inputs').Cell(strcat(['B' int2str(i+1)])).CellVariable.Value;
    inputUB(i)=f.Operations.Item('Inputs').Cell(strcat(['C' int2str(i+1)])).CellVariable.Value;
    initialUs(i)=f.Operations.Item('Inputs').Cell(strcat(['D' int2str(i+1)])).CellVariable.Value;
    parameter.inputUnits(i,1)=cellstr(f.Operations.Item('Inputs').Cell(strcat(['E' int2str(i+1)])).CellText);
end

lb=strcat('lb = [',num2str(inputLB'),']'';\n');
ub=strcat('ub = [',num2str(inputUB'),']'';\n');
u0=strcat('u0 = [',num2str(initialUs'),']'';\n');
%%
fName = 'NLP4MATLAB.m';         %# A file name
fid = fopen(fName,'w');         %# Open the file

%Create the opt
if fid ~= -1
    fprintf(fid,...
        strcat(...
        'function [u_opt,fval,exitflag] = NLP4MATLAB()\n\n',...
        'clc\n',...
        'clear all\n',...
        'close all\n\n',...
        'global h;\n',...
        'global f;\n',...
        'global hyCase;\n',...
        'h = actxserver(''UnisimDesign.Application'');\n',...
        'hyCase = h.Activedocument;\n',...
        'f = hyCase.Flowsheet;\n\n',...
        'par=[];\n'));
%     fprintf(fid,'lb(%d) = f.Operations.Item(''Inputs'').Cell(''B%d'').CellVariable.Value;\n',inputIndices2);
%     fprintf(fid,'ub(%d) = f.Operations.Item(''Inputs'').Cell(''C%d'').CellVariable.Value;\n',inputIndices2);
    fprintf(fid,'lb=zeros(1,%d);\nub=ones(1,%d);\n',numberOfInputs,numberOfInputs);
    fprintf(fid,'u0(%d) = scaleInputs(f.Operations.Item(''Inputs'').Cell(''D%d'').CellVariable.Value,lb,ub,1);\n',inputIndices2);
    fprintf(fid,...
        strcat(...
        'options = optimset(''TolFun'',10e-8,''TolCon'',1e-4,''Display'',''iter'',''Algorithm'',''interior-point'',''Diagnostics'',''on'', ''FinDiffType'',''central'',''ScaleProblem'',''obj-and-constr'',''FinDiffRelStep'',1e-2);\n',...
        'tic\n',...
        '[u_opt,fval,exitflag]=fmincon(@(u)objFun(u,par),u0,[],[],[],[],lb,ub,@(u)nonLinConFun(u,par),options);\n',...
        'toc\n',...
        'end\n\n'...
        ));
end
%%
if fid ~= -1
    fprintf(fid,...
        strcat(...
        'function y = objFun(u,par)\n',...
        'global f;\n',...
        'global hyCase;\n',...
        'hyCase.Solver.CanSolve=0;\n'...
        ));
    fprintf(fid,'lb(%d) = f.Operations.Item(''Inputs'').Cell(''B%d'').CellVariable.Value;\n',inputIndices2);
    fprintf(fid,'ub(%d) = f.Operations.Item(''Inputs'').Cell(''C%d'').CellVariable.Value;\n',inputIndices2);
    fprintf(fid,'input%dUnits = f.Operations.Item(''Inputs'').Cell(''E%d'').CellText;\n',inputIndices);
    fprintf(fid,'f.Operations.Item(''Inputs'').Cell(''A%d'').CellVariable.SetValue(deScaleInputs(u,lb,ub,%d),input%dUnits);\n',inputIndices3);

    fprintf(fid,...
        strcat(...        
        'hyCase.Solver.CanSolve=1;\n',...
        'y = f.Operations.Item(''Objective'').Cell(''A2'').CellValue;\n',...
        'end\n\n'...
        ));   
end

%% Read the Constraints spreadsheet
% Count the constraints
i=0;
while f.Operations.Item('Constraints').Cell(strcat(['A' int2str(i+2)])).CellVariable.IsKnown
    i=i+1;
end
numberOfInequalityConstraints=i;

i=0;
while f.Operations.Item('Constraints').Cell(strcat(['B' int2str(i+2)])).CellVariable.IsKnown
    i=i+1;
end
numberOfEqualityConstraints=i;

ineqConIndices=[2:1:numberOfInequalityConstraints+1;2:1:numberOfInequalityConstraints+1];
eqConIndices=[2:1:numberOfEqualityConstraints+1;2:1:numberOfEqualityConstraints+1];

if fid ~= -1
    fprintf(fid,...
        strcat(...
        'function [c,ceq] = nonLinConFun(u,par)\n',...
        'global f;\n',...
        'global hyCase;\n',...
        'hyCase.Solver.CanSolve=0;\n'...
        ));
    fprintf(fid,'lb(%d) = f.Operations.Item(''Inputs'').Cell(''B%d'').CellVariable.Value;\n',inputIndices2);
    fprintf(fid,'ub(%d) = f.Operations.Item(''Inputs'').Cell(''C%d'').CellVariable.Value;\n',inputIndices2);
    fprintf(fid,'input%dUnits = f.Operations.Item(''Inputs'').Cell(''E%d'').CellText;\n',inputIndices);
    fprintf(fid,'f.Operations.Item(''Inputs'').Cell(''A%d'').CellVariable.SetValue(deScaleInputs(u,lb,ub,%d),input%dUnits);\n',inputIndices3);

    fprintf(fid,'hyCase.Solver.CanSolve=1;\n');
    
    if numberOfInequalityConstraints>0
        fprintf(fid,'c(%d)=f.Operations.Item(''Constraints'').Cell(''A%d'').CellValue;\n',ineqConIndices);
    else
        fprintf(fid,'c=[];\n');
    end
    if numberOfEqualityConstraints>0
        fprintf(fid,'ceq(%d)=f.Operations.Item(''Constraints'').Cell(''B%d'').CellValue;\n',eqConIndices);
    else
        fprintf(fid,'ceq=[];\n');
    end
    fprintf(fid,'end\n\n');   
end
%%
if fid ~= -1
    fprintf(fid,...
        strcat(...
        'function y = scaleInputs(u,lb,ub,index)\n',...        
        'y=(u(index)-lb(index))/(ub(index)-lb(index));\n',...
        'end\n\n'...
        ));   
end

%%
if fid ~= -1
    fprintf(fid,...
        strcat(...
        'function y = deScaleInputs(u,lb,ub,index)\n',...        
        'y=lb(index)+u(index)*(ub(index)-lb(index));\n',...
        'end\n\n'...
        ));   
end

fclose(fid);
