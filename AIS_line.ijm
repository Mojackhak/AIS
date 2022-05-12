/*this program is used to get the intensity profile 
 * of the ais proteins along the axon 
 */

/*Usage:
 *confocal camera settings and the IHC are designed as below:
  Channel 1 = 405 PL→cPL; Channel 2 = 488 protein that you are intrested in ;
  Channel 3 = 555 PL→BLA; Channel 4 = 647  AnkG  
 *this Macro will output the results to excel (path and name is the same as the image)
 *see details in instructions.docx
 */

/*Author: Ankang Hu, Fudan University,
 * 		  18301050300@fudan.edu.cn
 * 2021/06
 */





/*
 ****************************************************************************************************
 */
 //vsi is the 10x image used to get the coordinates of neurons
#@ File (label="select the vsi file", style="file") vsi_path
 //nd2 is the 60x image used to get AIS profile of neurons
image_file = getDirectory ("select the nd2 file");
image_list = getFileList(image_file);
run("Brightness/Contrast...");
run("Channels Tool...");
vsi_title = import_vsi(vsi_path);
line_parameters = setaxis(vsi_title);
cycle = lengthOf(image_list);
loops = 0;
while (loops < cycle) {
	image_path = image_file+image_list[loops];	
	doloop(vsi_path, image_path, line_parameters);
	loops++;
}

//close all the windows
list_window = getList("window.titles");
list_image = getList("image.titles");
for (i = 0; i < lengthOf(list_window); i++) {
	close(list_window[i]);
}
for (i = 0; i < lengthOf(list_image); i++) {
	close(list_image[i]);
}


/*
 ***************************************************************************************************************
 */
 
function doloop(vsi_path, image_path, line_parameters) { 
/* function description:
 * do the loop until all the confocal image of this brain slice has been counted.
 */
 
	/*to import a confocal image and creat a z-stacked channel merge image
 * savename is the place to save the result excel
 */
	image_title=import_nd2(image_path);
	
	//set the start sheet to save the results in excel
	sheet_number = 0;
	sheet_number_BLA = 0;
	sheet_number_cPL = 0;
	ans = 1;
	
	//get the slices of the image
	selectWindow(image_title);
	slices = nSlices/4;
	
	//get the results of an ais in the way we want
	while (ans == 1) {
		num = get_results(vsi_path, image_path, line_parameters, slices);	
		//judge whether to continue and save the results
		if (num == 1) {
			//projection type
			projection_number = getNumber ("projection cPL = 1 \nprojection BLA = 2", 1);
			if (projection_number == 1) {
				sheet_number_cPL++;
				sheet_number = sheet_number_cPL;
			}
			else {
				sheet_number_BLA++;
				sheet_number = sheet_number_BLA;
			}
			save_results(image_path, sheet_number, projection_number);
			close_non_nd2();
		}
		if (num == 2) {
			close_non_nd2();
		}
		if (num == 3) {
			//projection type
			projection_number = getNumber ("projection cPL = 1 \nprojection BLA = 2", 1);
			if (projection_number == 1) {
				sheet_number_cPL++;
				sheet_number = sheet_number_cPL;
			}
			else {
				sheet_number_BLA++;
				sheet_number = sheet_number_BLA;
			}
			save_results(image_path, sheet_number, projection_number);
			ans = num;
			close_non_nd2();
		}
		if (num == 4) {
			ans = num;
			close_non_nd2();
		}
	}
	
	//close all the non .vsi windows
	close_non_vsi();
}


function get_results(vsi_path, image_path, line_parameters, slices) { 
/* function description:
 *  get the results of an ais in the way we want
 */
 //setTool("freeline");
 	selectWindow("MAX_" + image_title);
 	
	// make sure the line has been drawn.
	hasline("draw a line along the ais then click Ok");

	/*Correspondance between line ROI coordinates on 
	  an image (2D XY) and along its profile (1D P)
	*/
	run("Fit Spline", "straighten");
	getSelectionCoordinates(x_ais, y_ais);
	
	//turn to unstacked image to judge if all the ais has been included
	selectWindow(image_title);
	makeSelection("freeline", x_ais, y_ais);	
	waitForUser("click ok if finish checking");
	
	//judge whether to continue and save the results
	num = getNumber ("save and continue = 1 \nabandon and continue = 2 \nsave and quit = 3 \nabandon and quit = 4", 2);

	if ((num == 1) || (num == 3) ) {
		//make sure all channels have open 
		selectWindow(image_title);
		Stack.setActiveChannels("1111");
		Table.create("Results");
		//get the profile intensity of all slices
		getallpixels(image_title, slices);		
		//get the ais width
		width = ais_width(image_title);
        setResult("width", 0, width);
		// get the cell location
		selectWindow(vsi_title);
		waitForUser("click Ok only if when you have selected the cell");
		selectWindow(vsi_title);
		getSelectionCoordinates(xcell, ycell);
		point_info = getdistance(vsi_title, xcell, ycell, line_parameters);
		for (i = 0; i < (lengthOf(point_info)/2); i++) {
			setResult("distance_x", i, point_info[2*i]);
			setResult("distance_y", i, point_info[2*i+1]);
		}           	        
	}	
	return num;
}

function close_non_nd2() { 
/* function description
 *  close all the windows that don't end with ".nd2" or start with "Plo"
 *  ---only left the confocal image window, B&C and Channels
 */
	list_window = getList("window.titles");
	list_image=getList("image.titles");
	list=Array.concat(list_window,list_image);
	for (i=0; i<list.length; i++){
		winame = list[i];
		if (winame.substring(0,3) == "Plo"){
	     	close(winame);
		} else{
			if((lengthOf(winame) != 0) && !(endsWith(winame, ".nd2")) && (winame != "B&C") && (winame != "Channels") && !(matches(winame, ".*vsi.*")) ){
	     		close(winame);
			}
		}
	}
}


function import_nd2(image_path) { 
/* function description
 * to import a confocal image and creat a z-stacked channel merge image
 * return the name array (image and savename) for saving the results
 */
	run("Bio-Formats Importer", "open=["+image_path+"] color_mode=Default rois_import=[ROI manager] view=Hyperstack stack_order=XYCZT");
	image_title=getTitle();	
	run("Split Channels");
	run("Merge Channels...", "c1=[C1-"+image_title+"] c2=[C2-"+image_title+"] c3=[C3-"+image_title+"] c4=[C4-"+image_title+"] create");
	run("Subtract Background...", "rolling=50");
	run("Z Project...", "projection=[Max Intensity]");
	return image_title;
}


//save the "Results" table to the excel
function save_results(image_path, sheet_number, projection_number) { 
/* function description
 *  to save the "Results" table to the excel
 */
 	if(projection_number == 1){
		projection = "cPL";
	}
	else {
		projection = "BLA";
	}
	
	//get the filename (exclude the suffix)
	image_path_length=lengthOf(image_path)-4;
	filename = image_path.substring(0,image_path_length);
	savename_ais = filename + "_" + projection + ".xlsx";
	
	run("Read and Write Excel", "no_count_column file=["+savename_ais+"] sheet=AIS_" + sheet_number + " dataset_label=Project_PL-" + projection); 	
	close("Results");
}


function hasline(message) { 
/* function description
 * to make sure there is a line drawn by the user
 */
 	waitForUser(message);
	a=is("line");	
	while (a == 0) {
		waitForUser(message);
		a=is("line");
	}
}

function getallpixels(image_title, slices) { 
/* get the pixels intensity of all slices and channels along the profile 
*/
	//set varibles
	for (i = 1; i <= slices; i++) {
	    selectWindow(image_title);
		Stack.setPosition(2, i, 1);
		ypoints = getProfile();
		Stack.setPosition(4, i, 1);
		y_allponits = getProfile();
	    y_allponits = Array.concat(y_allponits,ypoints);
	    selectWindow("Results");
		Table.setColumn("slice_" + i + "", y_allponits);
	}
}

function ais_width(image_title) { 
/* get the calibrated distance of ais
 * pixel_length is determinaed by the image parameters
 */
	 //setTool("straight line");
 	selectWindow(image_title);
 	
	// make sure the line has been drawn.
	hasline("draw a line across the ais as the width");
	width = getValue("Length");	
	return width;
}


function close_non_vsi() { 
/* function description
 *  close all the windows that don't end with ".nd2" or start with "Plo"(only left the confocal image window) 
 *  also left B&C and Channels
 */
	list_window = getList("window.titles");
	list_image=getList("image.titles");
	list=Array.concat(list_window,list_image);
	for (i=0; i<list.length; i++){
		winame = list[i];
		if (winame.substring(0,3) == "Plo"){
	     	close(winame);
		} else{
			if((lengthOf(winame) != 0) && !(matches(winame, ".*vsi.*")) && (winame != "B&C") && (winame != "Channels")){
	     	close(winame);
			}
		}
	}
}


function setaxis(vsi_title) { 
/*In the coronal section, seting the median fissure of the brain slice as the y-axis
 * set the top of the median fissure of the brain slice as the original point
 */
	 //setTool("straight line");
 	selectWindow(vsi_title);
 	
	// make sure the line has been drawn.
	hasline("draw a line as the y axis");
	getSelectionCoordinates(x_ais, y_ais);
	original_x = x_ais[0];
	original_y = y_ais[0];
	last_x = x_ais[lengthOf(x_ais)-1];
	last_y = y_ais[lengthOf(y_ais)-1];

	//calculate the parameters of the line
	A = last_y - original_y;
	B = original_x - last_x;
	C = original_x*(original_y-last_y)+original_y*(last_x-original_x);

	line_parameters = newArray(5);
	line_parameters[0] = original_x;
	line_parameters[1] = original_y;
	line_parameters[2] = A;
	line_parameters[3] = B;
	line_parameters[4] = C;

	return line_parameters;

}


function getdistance(vsi_title, xpoints, ypoints, line_parameters) { 
/* function description
 *  get the x,y distance of the cyto
 */
 	original_x = line_parameters[0];
	original_y = line_parameters[1];
	A = line_parameters[2];
	B = line_parameters[3];
	C = line_parameters[4];
	point_info = newArray(2*lengthOf(xpoints));
	selectWindow(vsi_title);
	getPixelSize(unit, pixelWidth, pixelHeight);
	for (i = 0; i < lengthOf(xpoints); i++) {
		distance = sqrt(pow((xpoints[i] - original_x), 2) + pow((ypoints[i] - original_y), 2));
		distance_x = abs(A * xpoints[i] + B * ypoints[i] + C) / sqrt(pow(A, 2) + pow(B, 2));		
		distance_y = sqrt(pow(distance, 2) - pow(distance_x, 2));
		point_info[2*i] = distance_x * pixelWidth;
		point_info[2*i+1] = distance_y * pixelWidth;
	}
	return point_info;
}


function import_vsi(vsi_path) { 
/* function description
 * to import a confocal image and creat a z-stacked channel merge image
 * return the name array (image and savename) for saving the results
 */
	run("Bio-Formats Importer", "open=["+vsi_path+"]");
	vsi_title = getTitle();
	run("Split Channels");
	run("Merge Channels...", "c1=[C1-"+vsi_title+"] c2=[C2-"+vsi_title+"] c3=[C3-"+vsi_title+"] create");
	return vsi_title;
}