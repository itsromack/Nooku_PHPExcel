<?php
/**
 * Default Xls View
 *
 * @author      Romack Natividad <romack@wizmediateam.com>
 * @subpackage  Default
 */
class ComDefaultViewXls extends KViewTemplate
{
	var $_title;
	var $_filename;
	var $_headers;
	var $_description;

    /**
     * Initializes the config for the object
     *
     * Called from {@link __construct()} as a first step of object instantiation.
     *
     * @param 	object 	An optional KConfig object with configuration options
     * @return  void
     */
    protected function _initialize(KConfig $config)
    {
    	$config->append(array(
			'mimetype'	  => 'application/vnd.ms-excel;',
			'title' => 'Spreadsheet Title',
			'filename' => 'Spreadsheet-File',
			'headers' => null,
			'description' => 'Spreadsheet-Description'
       	));
       	$this->_title = $config->title;
       	$this->_headers = $config->headers;
       	$this->_filename = $config->filename;
       	$this->_description = $config->description;
    	
    	parent::_initialize($config);
    }
	
	/**
	 * Return the views output
	 * 
	 * This function will auto assign the model data to the view if the auto_assign
	 * property is set to TRUE. 
 	 *
	 * @return string 	The output of the view
	 */
	public function display()
	{
	    if(empty($this->output))
		{
	        $model = $this->getModel();
			
		    //Auto-assign the state to the view
		    $this->assign('state', $model->getState());
		
		    //Auto-assign the data from the model
		    if($this->_auto_assign)
		    {
			    //Get the view name
			    $name  = $this->getName();
		
			    //Assign the data of the model to the view
			    if(KInflector::isPlural($name))
			    {
				    $this->assign($name, 	$model->getList())
					     ->assign('total',	$model->getTotal());
			    }
			    else $this->assign($name, $model->getItem());
		    }
		}

		$user = JFactory::getUser();
		$generated_by = ($user->id > 0) ? $user->name : 'Guest';

		// Get and set the document properties
	    $document = &JFactory::getDocument();

	    $date = new JDate();
	    $format = '%Y-%m-%d';

	    $download_desc = JText::sprintf('DOWNLOAD DESC', 
	    	JText::_($this->_description), 
	    	$generated_by, 
	    	$date->toFormat($format), 
	    	$generated_by);

	    $document->setDescription($download_desc);
	    $document->setName(JText::_($this->_filename) . '-' . $date->toFormat($format));
	    $download_title = JText::sprintf('DOWNLOAD TITLE', JText::_($this->_title), $generated_by);
	    $document->setTitle($download_title);

	    // Get the PHPExcel object to set some properties
	    $phpexcel =& $document->getPhpExcelObj();
	    $phpexcel->getProperties()->setCreator($generated_by)->setLastModifiedBy($generated_by);
	    $phpexcel->setActiveSheetIndex(0);
	    $phpexcel->getActiveSheet()->getHeaderFooter()->setOddHeader('&C' . $download_desc);
	    $phpexcel->getActiveSheet()->setTitle(JText::_($this->_title));

	    $this->assign('headers', $this->_headers);
	    $this->assign('total_headers', count($this->_headers));

	    // Assign the PHPExcel object for use in the template
	    $this->assign('phpexcel', $phpexcel);

	    // variables below are used for cell locations
	    $this->assign('reset_col_char', 'A');
	    $this->assign('reset_col_index', 0);
	    $this->assign('reset_data_row_index', 2); # start at 2nd row because 1st row is for headers

	    // Cell Alignments
		$this->assign('align_right', array('alignment' => array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT)));
		$this->assign('align_center', array('alignment' => array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER)));

		// Font Styles
		$this->assign('font_bold', array('font' => array('bold' => true)));

		// Cell Borders
		$border_top = array('top' 	=> array('style' => PHPExcel_Style_Border::BORDER_THIN));
		$border_bottom = array('bottom' 	=> array('style' => PHPExcel_Style_Border::BORDER_THIN));
		$border_left = array('left' 	=> array('style' => PHPExcel_Style_Border::BORDER_THIN));
		$border_right = array('right' 	=> array('style' => PHPExcel_Style_Border::BORDER_THIN));
		$this->assign('border_top', $border_top);
		$this->assign('border_bottom', $border_top);
		$this->assign('border_left', $border_top);
		$this->assign('border_right', $border_top);
		$this->assign('border_box', array_merge($border_top, $border_bottom, $border_left, $border_right));
		
		return parent::display();
	}
}