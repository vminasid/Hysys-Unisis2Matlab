# Hysys/Unisim interfaced with Matlab's Optimization Toolbox
This script auto-generates the necessary MATLAB code in order to interface an existing Unisim/Hysis model with MATLAB's Optimization Toolbox. 
$$x^2$$
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
