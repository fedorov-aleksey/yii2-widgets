<?php
/**
 * Created by PhpStorm.
 * User: ais
 * Date: 08.12.16
 * Time: 16:47
 */

namespace fav\widgets;

use yii\base\Exception;
use yii\base\Widget;
use yii\base\InvalidConfigException;
use yii\base\InvalidParamException;


/**
 * The GridExport widget is used to export data in excel file.
 *
 * A basic usage looks like the following:
 *
 * ```php
 * <?= GridExport::widget([
 *     'dataProvider' => $dataProvider,
 *     'columns' => [
 *	    // first sheet
 *         [
 *		[
 *		    'attribute',
 *		    'label',
 *		    'options'=>[]
 *		],
 *		[
 *		    'attribute',
 *		    'label',
 *		    'options'=>[]
 *		],
 *		//....
 *	   ],
 *	    // second sheet
 *	   [
 *		[
 *		    'attribute',
 *		    'label',
 *		    'options'=>[]
 *		],
 *		[
 *		    'attribute',
 *		    'label',
 *		    'options'=>[]
 *		],
 *		//....
 *	   ],
 *     ],
 *	//....
 * ]) ?>
 * ```

*/

use yii\data\ArrayDataProvider;

class GridExport extends Widget
{
    
    /**
     * @var ArrayDataProvider
     */
    public $dataProvider;
    
    /**
     * @var array used to implement styles for table header. 
     * Should be specified for each sheet.
     * Can include next options:
     * - `background-color`
     * - `color` set the font color
     * - `multiline` set wrap text to cells
     * - `border` set thin border to cells
     * - `border-color`
     * - `font-family`
     * - `font-size`
    */
    
    public $headerRowOptions = [];
    
    /**
     * @var array used to implement styles for all table content. 
     * Should be specified for each sheet.
     * Can include next options:
     * - `background-color`
     * - `color` set the font color
     * - `multiline` set wrap text to cells
     * - `border` set thin border to cells
     * - `border-color`
     * - `font-family`
     * - `font-size`
    */
    public $contentOptions = [];
    
    /**
     * @var array used to implement styles for all table content. 
     * Should be specified for each sheet.
     * Can include next options:
     * - `background-color`
     * - `color` set the font color
     * - `multiline` set wrap text to cells
     * - `border` set thin border to cells
     * - `border-color`
     * - `font-family`
     * - `font-size`
     * - `row-renderer` set row styles according to defined rules
    */
    public $rowOptions = [];
    
    /**
     * @var array used to specify columns of the grid
     * Should be specified for each sheet.
     * [
     *		'attribute'=>..., //neseccary when $this->useAttributes is true
     *		'label'=>...,
     *		'options'=>[
     *		    ...
     *		]
     * ],
     * Can include next options:
     * - `background-color`
     * - `color` set the font color
     * - `multiline` set wrap text to cells
     * - `border` set thin border to cells
     * - `border-color`
     * - `font-family`
     * - `font-size`
     * - `col-width` set column width. Also can be `auto`
     * - `col-renderer` set column cells styles and values according to defined rules
    */
    
    public $columns = [];
    
    /**
     * @var string export filename
    */
    public $fileName = 'export.xls';
    
    /**
     * @var string export file format
    */
    public $fileFormat = 'Excel2007';
    
    /**
     * @var string export filename path. 
     * When defined export file name will be held here. Else php://output
    */
    public $savePath;
    
    /**
     * @var array sheet titles
     * Should be defined for each sheet. If undefined default titles 'Table 1, Table2 ... ' will be used
    */
    public $sheetTitles = [];
    
    /**
     * @var array used for downloading file
     * If true - file will be downloaded. If false - will be saved only
    */
    public $asAttachment = true;
    
    /**
     * @var array used to implement attributes in columns and models
     * If true - columns should inlude `attribute` option. And models should be ['attribute_name1'=>'value1','attribute_name2'=>'value2',...]
     * Else columns don`t need `attrubute` option and models should be ['value1','value2',...]
    */
    public $useAttributes = true;
    
    /**
     * @var PHPExcel used to construct excel file
    */
    protected $objExcel;
    
    /**
     * @var array functions which implemented styles to current sheet
    */
    protected $styleFunctions;
    
    /**
     * @var array data models of current sheet
    */
    protected $currentModels;
    
    /**
     * Initializes the grid export, phpexcel object and checking base params of the widget config.
     */
    public function init()
    {
        parent::init();
	$this->objExcel = new \PHPExcel();
	
	// checking only base params
	$this->checkConfig();
    }
    
    /**
     * Validates base params of the widget config
    */
    protected function checkConfig(){
	if(!$this->dataProvider){
	    throw new InvalidConfigException('\'dataProvider\' must be set');
	}
	else if(! ($this->dataProvider instanceof \yii\data\BaseDataProvider)){
	    throw new InvalidConfigException('\'dataProvider\' must be dataProvider class');
	}
	
	if(empty($this->columns)){
	    throw new InvalidConfigException('\'columns\' must be set');
	}
	
	if(empty($this->fileName)){
	    throw new InvalidConfigException('\'columns\' must be set');
	}
	
	if(empty($this->fileFormat)){
	    throw new InvalidConfigException('\'columns\' must be set');
	}
    }
    
    /**
     * Initializes styles functions in array `styleFunctions`
     */
    private function initStyleFunctions(){
	$this->styleFunctions = [];
	
	/**
	* Used to set background color.
	* Coords example 'A1:C1'
	* Color example '#FFFFFF' or 'FFFFFF'
	*/
	$this->styleFunctions['background-color'] = function($coords,$color){
	    
	    // set solid filling
	    $this->objExcel->getActiveSheet()
		    ->getStyle($coords)
		    ->getFill()
		    ->setFillType(\PHPExcel_Style_Fill::FILL_SOLID);
	    
	    // set filling color
	    $this->objExcel->getActiveSheet()
		    ->getStyle($coords)
		    ->getFill()
		    ->getStartColor()
		    ->setRGB(str_replace('#', '', $color));
	};
	
	/**
	* Used to set font color.
	* Coords example 'A1:C1'
	* Color example '#FFFFFF' or 'FFFFFF'
	*/
	$this->styleFunctions['color'] = function($coords,$color){
	    $this->objExcel->getActiveSheet()
		    ->getStyle($coords)
		    ->getFont()
		    ->getColor()
		    ->setRGB(str_replace('#', '', $color));
	};
	
	/**
	* Used to make cells multiline
	* Coords example 'A1:C1'
	*/
	$this->styleFunctions['multiline'] = function($coords,$value=true){
	    $this->objExcel->getActiveSheet()
		    ->getStyle($coords)
		    ->getAlignment()
		    ->setWrapText($value);
	};
	
	
	/**
	* Used to set thin or none border to cells
	* Coords example 'A1:C1'
	*/
	$this->styleFunctions['border'] = function($coords,$is_bordered=false){
	    $is_bordered = ($is_bordered) 
		    ? \PHPExcel_Style_Border::BORDER_THIN 
		    : \PHPExcel_Style_Border::BORDER_NONE;
	    
	    $this->objExcel->getActiveSheet()
		    ->getStyle($coords)
		    ->getBorders()
		    ->getAllBorders()
		    ->setBorderStyle($is_bordered);
	};
	
	/**
	* Used to set border color
	* Coords example 'A1:C1'
	* Color example '#FFFFFF' or 'FFFFFF'
	*/
	$this->styleFunctions['border-color'] = function($coords,$color){
	    $this->objExcel->getActiveSheet()
		    ->getStyle($coords)
		    ->getBorders()
		    ->getAllBorders()
		    ->getColor()
		    ->setRGB(str_replace('#', '', $color));
	};
	
	/**
	* Used to set column styles according to defined rules
	* Rules example
	* function($value){
	*	if(...){
	*	    return [
	*		'value'=>...,
	*		'options'=>[
	*		    ....
	*		]
	*	    ];
	*	}
	*	else{
	*	    return ...;
	*	}
	* }
	* Options should include:
	* - `column-index` - number of column (from 0 to ...)
	* 
	*/
	$this->styleFunctions['col-renderer'] = function($options,$rules){
	    $models = $this->currentModels;
	    
	    foreach($models as $key=>$model){
		// calc coords of column cells
		$coords = $this->num2alpha($options['column-index']).($key+2);
		
		// get result from callback function
		$result = $rules($model[
		    (
			isset($this->columns[$options['column-index']]['attrubute'])
			? $this->columns[$options['column-index']]['attrubute']
			: $options['column-index']
		    )
		]);
		
		// if result is array - implement all returned options
		if(is_array($result)){
		    $this->objExcel
			->getActiveSheet()
			->setCellValue(
				$coords,
				$result['value']
			);
		    
		    $this->provideStylesArray($coords,$result['options']);
		}
		// if value - set it to cells
		else{
		    $this->objExcel
		    ->getActiveSheet()
		    ->setCellValue(
			    $coords,
			    $result
		    );
		}
	    }
	};
	
	/**
	* Used to set column width. Also can be `auto`
	* Column example 'A','B',...
	* Width example 20,30,.....,'auto'
	*/
	$this->styleFunctions['col-width'] = function($column,$width){
	    
	    if($width == 'auto'){
		$this->objExcel->getActiveSheet()
			->getColumnDimension($column)
			->setAutoSize(true);
	    }
	    else{
		$this->objExcel->getActiveSheet()
			->getColumnDimension($column)
			->setWidth($width);
	    }
	};
	
	/**
	* Used to set row styles according to defined rules
	* Rules example
	* function($value){
	*	if(...){
	*	    return [
	*		'value'=>...,
	*		'options'=>[
	*		    ....
	*		]
	*	    ];
	*	}
	*	else{
	*	    return ...;
	*	}
	* }
	* Options should include:
	* - `row-index` - number of row (from 0 to ...)
	* 
	*/
	$this->styleFunctions['row-renderer'] = function($options,$rules){
	    $models = $this->currentModels;
	    
	    // get result from callback function
	    $result = $rules($this->currentModels[$options['row-index']]);
	    // calc cells coords
	    $rownum = ($options['row-index']+2);
	    $coords = 'A'.$rownum.':'.$this->num2alpha(count($this->columns)-1).$rownum;
	    
	    //implement all returned options
	    $this->provideStylesArray($coords,$result['options']);
	};
	
	/**
	* Used to set font size
	* Coords example 'A1:C1'
	* Size  - from 1 to ...
	*/
	$this->styleFunctions['font-size'] = function($coords,$size){
	    $this->objExcel->getActiveSheet()->getStyle($coords)->getFont()->setSize($size);
	};
	
	/**
	* Used to set font name
	* Coords example 'A1:C1'
	* Font - name of font. Example - 'Times New Roman'
	*/
	$this->styleFunctions['font-family'] = function($coords,$font){
	    $this->objExcel->getActiveSheet()->getStyle($coords)->getFont()->setName($font);
	};
    }
    
    /**
     * Converts numeric index to excel column index
     * @param integer $n numeric index
     * @return string excel column index
     */
    protected function num2alpha($n)
    {
	for($r = ""; $n >= 0; $n = intval($n / 26) - 1)
	    $r = chr($n%26 + 0x41) . $r;
	return $r;
    }

    /**
     * Creates sheet, fill it and implenents styles
     */
    protected function initSheets(){
	
	// get numeric keys from data provider to get sheets index array
	$sheets = $this->dataProvider->getKeys();
	
	if(!empty($sheets)){
	    // if no sheet titles defined use default sheet titles 'Table ...'
	    if(empty($this->sheetTitles)){
		foreach($sheets as $key=>$sheet){
		    $this->sheetTitles[$key] = 'Table '.($key+1);
		}
	    }
	    else if(count($sheets)!=count($this->sheetTitles)){
		throw new InvalidConfigException('Size of \'sheetTitles\' does not match count of sheets of \'dataProvider\'');
	    }

	    foreach($sheets as $key=>$sheet){
		// use default sheet for first list, create new sheet for others
		if($key){
		    $this->objExcel->createSheet($key);
		}
		// setting current active sheet
		$this->objExcel->setActiveSheetIndex($key);
		// set sheet title
		$this->objExcel->getActiveSheet()->setTitle($this->sheetTitles[$key]);
		
		// fill sheet data, implement styles
		$this->prepareSheet($key);
	    }
	    
	    // set first sheet is active
	    $this->objExcel->setActiveSheetIndex(0);
	}
    }
    
    /**
     * Fills sheet data, implement styles
     * @param integer $index sheet index
     */
    protected function prepareSheet($index){
	
	if(!isset($this->columns[$index])){
	    throw new InvalidConfigException('No columns defined for sheet '.$index);
	}
	
	// fill head row
	$this->initHeadRow($index);
	
	// fill cells
	$rows = $this->initCells($index);
	
	// implement styles
	$this->implemetContentStyles($index,$rows);
	$this->implementColumsStyles($index,$rows);
	$this->implementRowStyles($index);
	$this->implementHeaderStyles($index);
    }
    
    /**
     * Fills head row
     * @param integer $index sheet index
     */
    protected function initHeadRow($index){
	
	$headers = array_column($this->columns[$index], 'label');
	
	foreach($headers as $key=>$header){
	    $this->objExcel
		    ->getActiveSheet()
		    ->setCellValue(
			    $this->num2alpha($key).'1',
			    $header
		    );
	}
    }
    
    /**
     * Fills cells
     * @param integer $index sheet index
     * @return integer number of rows (with header)
     */
    protected function initCells($index){
	
	
	
	if(!isset($this->dataProvider->getModels()[$index])){
	    throw new InvalidConfigException('No data for Sheet '.$index.' defined in dataProvider');
	}
	
	$this->currentModels = $this->dataProvider->getModels()[$index];
	
	// get models attrubutes if necessary or init default numbers
	$attributes = ($this->useAttributes)
		    ? array_column($this->columns[$index], 'attribute')
		    : range(0, count($this->columns[$index])-1);
	
	if(empty($attributes)){
	    throw new InvalidConfigException('No attributes defined in \'columns\'');   
	}
	
	$rows = 1; // rows counter
	foreach($this->currentModels as $model){
	    $rows++;
	    // set value for each cell
	    foreach($attributes as $key=>$attribute){
		
		$this->objExcel
		    ->getActiveSheet()
		    ->setCellValue(
			    $this->num2alpha($key).$rows,
			    $model[$attribute]
		    );
	    }
	}
	return $rows;
    }
    
    /**
     * Saves file
     */
    protected function writeFile(){

	// init wrighter
	$objectwriter = \PHPExcel_IOFactory::createWriter(
		$this->objExcel
		, $this->fileFormat
	);
	
	$path = 'php://output';
	
	// if savePath defined implement it
	if (isset($this->savePath) && $this->savePath != null) {
	    $path = $this->savePath . '/' . $this->fileName();
	}
	
	$objectwriter->save($path);
	exit();
    }
    
    /**
     * Implements array of styles
     * @param string $coords area to process
     * @param array $styles collection of styles to implement
     */
    private function provideStylesArray($coords,$styles){
	// exclude difficult styles
	$skipStyles = ['col-width','col-renderer','row-renderer'];
	
	if(!empty($styles)){
	    foreach($styles as $key=>$value){
		if(
			isset($this->styleFunctions[$key])
			&& !in_array($key, $skipStyles)
		){
		    $this->styleFunctions[$key]($coords,$value);
		}
		else{
		    if(!isset($this->styleFunctions[$key])){
			throw new InvalidParamException('Style function \''.$key.'\' is undefined');
		    }
		}
	    }
	}
    }
    
    /**
     * Implements array of styles to all table
     * @param integer $index sheet index
     * @param integer $rows rows amount (with header)
     */
    protected function implemetContentStyles($index,$rows){
	$coords = 'A1:'
		.$this->num2alpha(count($this->columns[$index])-1)
		.$rows;
	
	$options = (isset($this->contentOptions[$index])) 
		    ? $this->contentOptions[$index]
		    : [];
	
	$this->provideStylesArray(
		$coords, 
		$options
	);
    }
    
    /**
     * Implements array of styles to data row (skips header)
     * @param integer $index sheet index
     */
    protected function implementRowStyles($index){
	
	$options = isset($this->rowOptions[$index])
		? $this->rowOptions[$index]
		: [];
	
	foreach($this->currentModels as $key=>$model){
	    
	    $coords = 'A'.($key+2).':'.$this->num2alpha(count($this->columns[$index])-1).($key+2);
	    
	    $this->provideStylesArray(
		$coords, 
		$options
	    );
	    
	    if(isset($options['row-renderer'])){
		$this->styleFunctions['row-renderer'](
		    [
			'row-index'=>$key,
		    ],
		    $options['row-renderer']
		);
	    }
	}
    }
    
    /**
     * Implements array of styles to data column (skips header)
     * @param integer $index sheet index
     * @param integer $rows rows amount (with header)
     */
    protected function implementColumsStyles($index,$rows){
	
	foreach($this->columns[$index] as $key=>$column){
	    if(isset($column['options'])){
		$col_index = $this->num2alpha($key);
		$coords = $col_index.'2:'.$col_index.$rows;
		
		$this->provideStylesArray(
			$coords, 
			$column['options']
		);
		
		if(isset($column['options']['col-width'])){
		    $this->styleFunctions['col-width'](
			    $col_index,
			    $column['options']['col-width']
		    );
		}
		
		if(isset($column['options']['col-renderer'])){
		    $this->styleFunctions['col-renderer'](
			    [
				'column-index'=>$key,
			    ],
			    $column['options']['col-renderer']
		    );
		}
	    }
	}
    }
    
    /**
     * Implements array of styles to header row (only)
     * @param integer $index sheet index
     */
    protected function implementHeaderStyles($index){
	$coords = 'A1:'
		.$this->num2alpha(count($this->columns[$index])-1)
		.'1';
	
	$options = (isset($this->headerRowOptions[$index])) ? 
		    $this->headerRowOptions[$index]
		    : [];
	
	$this->provideStylesArray($coords, $options);
    }
    
    /**
     * Set headers to download file
     */
    public function setHeaders()
    {
	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	header('Content-Disposition: attachment;filename="' . $this->fileName .'"');
	header('Cache-Control: max-age=0');
    }
    
    /**
     * Runs the widget.
     */
    public function run(){
	
	$this->initStyleFunctions();
	$this->initSheets();
	
	if ($this->asAttachment) {
	    $this->setHeaders();
	}
	
	//export file
	return $this->writeFile();
    }
}