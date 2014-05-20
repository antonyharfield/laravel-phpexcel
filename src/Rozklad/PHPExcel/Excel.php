<?php 
namespace Rozklad\PHPExcel;

use App;
use Config;

/**
 * A Laravel wrapper for PHPEXcel
 *
 * @version 0.1.0alpha
 * @package laravel-phpexcel
 * @author rozklad <jan.rozklad@gmail.com>
 */
class Excel {

	protected $phpexcel;

	public function __construct() 
	{
		$vendor_file = base_path() . '/vendor/phpoffice/phpexcel/Classes/PHPExcel.php';

		if ( file_exists( $vendor_file ) )
			require_once $vendor_file;
		else
			App::abort('500', "$vendor_file cannot be loaded, please run composer update to get the required phpoffice/phpexcel package");
	
		return $this->phpexcel;
	}

	protected function init()
	{
        $this->phpexcel = new \PHPExcel();

        // Set document properties
		$this->phpexcel->getProperties()->setCreator( Config::get('laravel-phpexcel::config.creator') )
			->setLastModifiedBy( Config::get('laravel-phpexcel::config.lastModifiedBy') )
			->setTitle( 'Untitled' )
			->setSubject( 'Subject' )
			->setDescription( 'Description' )
			->setKeywords( 'php keyword laravel excel' )
			->setCategory( 'Category' );
    }

    /**
     * [excel2Array description]
     * 
     * @param  [type] $filepath [description]
     * @param  array  $result   [description]
     * @return [type]           [description]
     */
    public function excel2Array( $filepath = null, $result = array() )
    {
    	if ( !file_exists( $filepath ) ) {
    		App::abort('500', "Error loading file ".$filepath.": File does not exist");
    	}

    	// Read your Excel workbook
		try {
		    $inputFileType = \PHPExcel_IOFactory::identify($filepath);
		    $objReader = \PHPExcel_IOFactory::createReader($inputFileType);
		    $objPHPExcel = $objReader->load($filepath);
		} catch(Exception $e) {
			App::abort('500', "Error loading file ".pathinfo($filepath,PATHINFO_BASENAME).": ".$e->getMessage());
		}
        
        $i = 0;
        // Loop through each worksheet
        foreach ($objPHPExcel->getWorksheetIterator() as $sheet) {
            // Get worksheet dimensions
            $highestRow = $sheet->getHighestRow(); 
            $highestColumn = $sheet->getHighestColumn();
            
            // Loop through each row of the worksheet in turn
            for ($row = 1; $row <= $highestRow; $row++){ 
                // Read a row of data into an array
                $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row,
                    NULL,
                    TRUE,
                    FALSE);
                $result[$i][] = $rowData;
            }
            $i++;
        }

		return $result;
    }

    /**
     * Create excel from array
     *
     * @uses PHPExcel\PHPExcel_Worksheet::fromArray()
     * @param   array   $source                 Source array
	 * @param   mixed   $nullValue              Value in source array that stands for blank cell
	 * @param   string  $startCell              Insert array starting from this cell address as the top left coordinate
     * @param   boolean $strictNullComparison   Apply strict comparison when testing for null values in the array
     */
	public function fromArray( $source = array(), $nullValue = NULL, $startCell = 'A1', $strictNullComparison = false ) 
	{
		$this->init();

		// Split $startCell to character and number
		$analyze = preg_split('/(?<=\d)(?=[a-z])|(?<=[a-z])(?=\d)/i', $startCell);
		$startCol = $analyze[0];
		$startRow = $analyze[1];

		$rowNumber = $startRow;
		foreach( $source as $row ) {
		    $this->phpexcel->getActiveSheet()->fromArray($row,NULL,$startCol.$rowNumber++);
		}

		return $this;
	}

	public function save( $filepath )
	{

		$ext = pathinfo($filepath, PATHINFO_EXTENSION);

		switch( $ext ) {

			case 'xlsx':
				$writerType = 'Excel2007';
			break;

			case 'xls':
			default:
				$writerType = 'Excel5';
			break;
		}

		
		try {
			$objWriter = \PHPExcel_IOFactory::createWriter($this->phpexcel, $writerType);
			$objWriter->save($filepath);
		} catch(Exception $e) {
			App::abort('500', "Error writing file ".pathinfo($filepath,PATHINFO_BASENAME).": ".$e->getMessage());
		}

		return $this;

	}

}