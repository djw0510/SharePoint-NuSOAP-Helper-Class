<?php
// Author: Danny Wenner
// This class is intended to simplify all other SharePoint integrations by processing input and output in a more friendly manner
// This file may be used on multiple projects, but the only line that should vary should be the inclusion path for the NuSOAP library
require_once(__DIR__.'/../hidden-connection-info/site-specific.php');

class sharepoint {
    public $spRoot;
    public $spSiteRoot;
    public $spListsWS;
    public $spUser;
    public $spPW;

    public $spService;
    public $spListName;
    public $camlQuery;
    public $camlWhere;
    public $camlOrderBy;
    public $camlOrderByAsc;
    public $camlViewFields;
    public $camlLimit;
    public $isRecursive;
    public $includeAttachmentUrls;
    public $camlUpdate;
    public $autoExecute;
    public $resultsArr;
    public $clientError;
    public $clientResponse;

    public function __construct(){
        global $spRoot, $spSiteRoot, $spListsWS, $spUser, $spPW; // set in hidden-connection-info/site-specific.php

        // arguments used when instantiating class using an array of variableName=>variableValue
        $arguments = func_get_args();
        if(!empty($arguments))
            foreach($arguments[0] as $k=>$v)
                if(property_exists($this, $k))
                    $this->{$k} = $v;

        // support to modify these variables to connect to other lists if desired
        if(empty($this->spRoot))
            $this->spRoot = $spRoot;
        if(empty($this->spSiteRoot))
            $this->spSiteRoot = $spSiteRoot;
        if(empty($this->spListsWS))
            $this->spListsWS = $spListsWS;
        if(empty($this->spUser))
            $this->spUser = $spUser;
        if(empty($this->spPW))
            $this->spPW = $spPW;

        // establish SOAP service provided we have the required elements
        if(!empty($this->spListsWS) && !empty($this->spUser) && !empty($this->spPW)){
            ini_set('error_reporting', 0); // silence errors in nusoap library
            require_once(__DIR__.'/nusoap/nusoap.php');
            $this->client = new nusoap_client($this->spListsWS, true);
            $this->client->setCredentials($this->spUser, $this->spPW, 'ntlm');
        } else {
            echo 'SharePoint credentials are undefined.';
            die();
        }

        // defaults
        if(empty($this->spService) && empty($this->camlUpdate))
            $this->spService = 'GetListItems';
        elseif(!empty($this->camlUpdate))
            $this->spService = 'UpdateListItems';

        if($this->camlOrderByAsc !== false)
            $this->camlOrderByAsc = true;

        if(empty($this->autoExecute))
            $this->autoExecute = false;
        // end defaults

        // if $autoExecute is true and $spService is set then automatically build and execute query
        if(!empty($this->autoExecute) && !empty($this->spService))
            $this->buildQuery();
    }

    function buildQuery(){
        $this->camlQuery = "<".$this->spService." xmlns='http://schemas.microsoft.com/sharepoint/soap/'>
                                <rowLimit>".(!empty($this->camlLimit) && is_int($this->camlLimit) ? $this->camlLimit : 0)."</rowLimit>
                                <listName>".$this->spListName."</listName>";
        if(!empty($this->isRecursive) || !empty($this->includeAttachmentUrls)){
            $this->camlQuery.= "<queryOptions>
                                    <QueryOptions>";

            if(!empty($this->isRecursive))
                $this->camlQuery.= "<ViewAttributes Scope='RecursiveAll'/>";
            if(!empty($this->includeAttachmentUrls))
                $this->camlQuery.= "<IncludeAttachmentUrls>TRUE</IncludeAttachmentUrls>";

            $this->camlQuery.= "</QueryOptions>
                            </queryOptions>";
        }
        if($this->spService == 'GetListItems'){
            $this->camlQuery.= "<query>
                                    <Query>";
            if(!empty($this->camlWhere))
                $this->camlQuery.= "<Where>".$this->camlWhere."</Where>";
            if(!empty($this->camlOrderBy) && is_array($this->camlOrderBy) && count($this->camlOrderBy) > 0){
                $this->camlQuery.= "<OrderBy>";
                foreach($this->camlOrderBy as $orderBy){
                    $orderBy = $this->formatFieldName($orderBy);
                    $orderByAsc = $this->camlOrderByAsc === true ? 'TRUE' : 'FALSE';

                    $this->camlQuery.= "<FieldRef Name='".$orderBy."' Ascending='".$orderByAsc."'/>";
                }
                $this->camlQuery.= "</OrderBy>";
            }
            $this->camlQuery.= "</Query>
                            </query>";
            if(!empty($this->camlViewFields) && is_array($this->camlViewFields) && count($this->camlViewFields) > 0){
                $this->camlQuery.= "<viewFields>";
                    $this->camlQuery.= "<ViewFields>";
                    foreach($this->camlViewFields as $field){
                        $field = $this->formatFieldName($field);
                        $this->camlQuery.= "<FieldRef Name='".$field."' />";
                    }
                    $this->camlQuery.= "</ViewFields>";
                $this->camlQuery.= "</viewFields>";
            } else {
                // retrieve all fields
                $this->camlQuery.= "<viewFields>";
                    $this->camlQuery.= "<ViewFields/>";
                $this->camlQuery.= "</viewFields>";
            }
        }
        elseif(!empty($this->camlUpdate)){
            // if your update is not executing properly make sure the field names in your CAML reflect the replacements made with formatFieldName()
            $this->camlQuery.= "<updates>".$this->camlUpdate."</updates>";
        }
        $this->camlQuery.= "</".$this->spService.">";

        // if $autoExecute is true then automatically execute query
        if(!empty($this->autoExecute))
            $this->executeQuery();
    }

    function executeQuery(){
        $return = $this->client->call($this->spService, $this->camlQuery);

        $clientError = $this->client->getError();
        $clientResponse = $this->client->response;
        unset($this->client);
        if(!empty($clientError)){
            $this->logError("SP Query Error: ".(!empty($this->spListName) ? $this->spListName : 'Unknown')." List | $clientError".(!empty($return['detail']['errorstring']) ? ' | '.$return['detail']['errorstring'] : '')." | Response: $clientResponse");
        } else {
            if(isset($return['GetListItemsResult']['listitems']['data']['row'])){
                $dataArr = $return['GetListItemsResult']['listitems']['data']['row'];

                if(!empty($this->camlViewFields) && is_array($this->camlViewFields) && count($this->camlViewFields) > 0){
                    // scrub data before saving - reduces file size dramatically
                    $filterFields = array();
                    if(isset($this->camlViewFields)){
                        foreach($this->camlViewFields as $field){
                            $field = $this->formatFieldName($field);
                            array_push($filterFields, '!ows_'.$field);
                        }
                    }

                    $dataArr_filtered = array();
                    if($return['GetListItemsResult']['listitems']['data']['!ItemCount'] == 1){
                        $tempArr = array();
                        for($x=0; $x<count($filterFields); $x++){
                            if(array_key_exists($filterFields[$x], $dataArr) && !empty($dataArr[$filterFields[$x]]))
                                $tempArr[$this->formatFieldName($filterFields[$x], 'clean')] = $dataArr[$filterFields[$x]];
                        }
                        array_push($dataArr_filtered, $tempArr);
                    }
                    elseif($return['GetListItemsResult']['listitems']['data']['!ItemCount'] > 1){
                        foreach($dataArr as $k=>$v){
                            $tempArr = array();
                            for($x=0; $x<count($filterFields); $x++){
                                for($y=0; $y<count($v); $y++){
                                    if(array_key_exists($filterFields[$x], $v) && !empty($v[$filterFields[$x]])){
                                        $tempArr[$this->formatFieldName($filterFields[$x], 'clean')] = $v[$filterFields[$x]];
                                        break;
                                    }
                                }
                            }
                            array_push($dataArr_filtered, $tempArr);
                        }
                    }

                    $this->resultsArr = $dataArr_filtered;
                } else {
                    // if no specific fields were requested then return all fields
                    $this->resultsArr = array();
                    foreach($dataArr as $data){
                        if(count($data) == 1){
                            foreach($dataArr as $k=>$v)
                                $tempArr[$this->formatFieldName($k, 'clean')] = $v;

                            $this->resultsArr[0] = $tempArr;
                        }
                        elseif(count($data) > 1){
                            $tempArr = array();
                            foreach($data as $k=>$v)
                                $tempArr[$this->formatFieldName($k, 'clean')] = $v;

                            array_push($this->resultsArr, $tempArr);
                        }
                    }
                }
            }
        }
    }

    // function to download files from SharePoint to local server
    function downloadFile($sourceFile=null, $downloadDestination=null){ // download file from SharePoint to specified destination
        // if $sourceFile does not contain an extension then assume it is a directory
        if(!strstr(basename($sourceFile), '.')){
            $this->logError('A directory was specified instead of a file.');
            die("A directory was specified instead of a file.");
        }

        if(!empty($sourceFile) && !empty($downloadDestination)){
            // if file is not from external site prepend with SharePoint path
            if(!stristr($sourceFile, '://')){
                if(substr($sourceFile, 0, 1) == '/')
                    $sourceFile = $this->spRoot.substr($sourceFile, 1, strlen($sourceFile)+1);
                else
                    $sourceFile = $this->spRoot.$sourceFile;
            }

            // format file path in order to properly retrieve files containing special characters from SharePoint
            $sourceFile = str_replace(basename($sourceFile), rawurlencode(basename($sourceFile)), $sourceFile); // rawurlencode() the filename without modifying the rest of the url
            $sourceFile = str_replace(' ', '%20', $sourceFile); // replace any spaces with rawurlencode() equivalent

            // if file already exists then delete it so it can be re-downloaded
            if(file_exists($downloadDestination))
                unlink($downloadDestination);

            // create file to output data as it downloads
            $fp = fopen($downloadDestination, 'w');

            // load file data using cURL
            $ch = curl_init();
            $optionsArr = array(
                CURLOPT_URL => $sourceFile,
                CURLOPT_RETURNTRANSFER => 1,
                CURLOPT_FOLLOWLOCATION => 1,
                CURLOPT_ENCODING => 'gzip,deflate',
                CURLOPT_FILE => $fp
            );
            // if $sourcePath contains SharePoint reference then set additional options // allows us to load external, non-SharePoint files without sending account data
            if(stristr($sourceFile, $this->spRoot)){
                $optionsArr[CURLOPT_HTTPAUTH] = CURLAUTH_NTLM;
                $optionsArr[CURLOPT_USERPWD] = $this->spUser.":".$this->spPW;
            }
            // set options for cURL transfer and execute
            curl_setopt_array($ch, $optionsArr);
            curl_exec($ch);

            // close output file
            fclose($fp);

            if(curl_error($ch)){
                echo logError('cURL error: '.curl_error($ch));
                unlink($downloadDestination); // delete the useless empty file that was created in the event that cURL fails
            }

            curl_close($ch);
        } else {
            echo 'File and/or destination are missing.';
            die();
        }
    }

    // function to upload files to SharePoint from local server
    function uploadFile($filePath, $library='Media', $folder=null){
        // replace any spaces with rawurlencode() equivalent
        if(!empty($library) && stristr($library, ' '))
            $library = str_replace(' ', '%20', $library);
        if(!empty($folder) && stristr($folder, ' '))
            $folder = str_replace(' ', '%20', $folder);

        if(file_exists($filePath)){
            $fileName = end(explode('/', $filePath));
            $data = file_get_contents($filePath);

            $ch = curl_init();
            $options = array(
                CURLOPT_URL => $this->spSiteRoot.$library.'/'.(!empty($folder) ? $folder.'/' : '').$fileName,
                CURLOPT_HTTPAUTH => CURLAUTH_NTLM,
                CURLOPT_USERPWD => $this->spUser.":".$this->spPW,
                CURLOPT_RETURNTRANSFER => 1,
                CURLOPT_CUSTOMREQUEST => "PUT",
                CURLOPT_POSTFIELDS => $data
            );
            curl_setopt_array($ch, $options);
            $ch_result = curl_exec($ch);
            $ch_api_err = curl_errno($ch);
            $ch_err_msg = curl_error($ch);
            curl_close($ch);

            //var_dump($ch_result);
            //var_dump($ch_api_err);
            //var_dump($ch_err_msg);
        } else {
            return false;
        }
    }

    // function to download contents of an entire SharePoint library or subdirectory to local server
    function downloadLibrary($listName=null, $listDirectory=null, $downloadDestination=null){
        /*if(!empty($downloadDestination)){
            if(!is_dir($downloadDestination))
                mkdir($downloadDestination, 0775);
            if(is_dir($downloadDestination) && substr(decoct(fileperms($downloadDestination)),2) != '775'){
                chmod($downloadDestination, 0775);
                $this->logError("Cache folder has insufficient permissions to perform caching.");
            }
        }*/

        // reset class for new query
        $this->resetClassObject();

        if(!empty($listName)){
            $this->spListName = $listName;
            $this->camlViewFields = array("FileRef", "Modified");
            $this->camlOrderBy = array("FileRef"); // will organize files alphabetically, directories first, followed by files
            $this->isRecursive = true;
            $this->autoExecute = true;

            // if specified, only get contents of specified directory
            if(!empty($listDirectory)){
                // make sure that specified directory ends with a forward-slash
                $listDirectory = rtrim($listDirectory, '/').'/';

                $this->camlWhere = "<Contains>
                                        <FieldRef Name='FileDirRef'/>
                                        <Value Type='Text'>".$listDirectory."</Value>
                                    </Contains>";
            }

            // reconstruct to set defaults and build/execute query
            $this->__construct();

            if(!empty($this->resultsArr)){
                for($i=0; $i<count($this->resultsArr); $i++){
                    // remove ID prefix from FileRef
                    $this->resultsArr[$i]['FileRef'] = end(explode(';#', $this->resultsArr[$i]['FileRef']));

                    // create necessary directory structure and download file(s)
                    $objectPath = __DIR__.'/..'.$downloadDestination.substr(end(explode($listName, $this->resultsArr[$i]['FileRef'])), 1); // determine path to object on server
                    if(strstr(end(explode('/', $this->resultsArr[$i]['FileRef'])), '.')){
                        // check if necessary directory structure exists
                        $dir = substr($objectPath, 0, strrpos($objectPath, '/')).'/';
                        if(!is_dir($dir)){
                            // make sure parent directory exists
                            $parentDir = substr($dir, 0, strrpos(rtrim($dir, '/'), '/')).'/';
                            if(!is_dir($parentDir))
                                mkdir($parentDir);

                            mkdir($dir);
                        }

                        // download file
                        $this->downloadFile($this->resultsArr[$i]['FileRef'], $objectPath);
                    }
                }

                /*echo "<pre>";
                var_dump($this->resultsArr);
                echo "</pre>";*/
            }
        }
    }

    // used to translate certain characters to the equivalents that SharePoint uses
    function formatFieldName($string, $action=null){
        $findReplaceArr = array(
            ' ' => '_x0020_',
            '-' => '_x002d_'
        );

        // optional // clean up fieldName // typically used when returning data rather than when querying
        if(!empty($action) && $action == 'clean'){
            $findReplaceArr = array_flip($findReplaceArr);
            $string = str_replace('!ows_', '', $string);
        }

        foreach($findReplaceArr as $k=>$v)
            $string = str_replace($k, $v, $string);

        return $string;
    }

    // used to reset class object
    function resetClassObject(){
        foreach($this as $k=>$v){
            if($k != 'spRoot' && $k != 'spSiteRoot' && $k != 'spListsWS' && $k != 'spUser' && $k != 'spPW')
                $this->$k = null;
        }
    }

    // to make it easier to report errors using a custom function if it is available
    function logError($errorMsg=NULL, $emailTo=NULL, $emailSubject=NULL){
        if(function_exists('logError'))
            logError($errorMsg, $emailTo, $emailSubject);
        else
            error_log($errorMsg);
    }
}
?>