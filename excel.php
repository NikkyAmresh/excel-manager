<?php

#namespace Aspera\Spreadsheet\XLSX;

#use RuntimeException;
#use SplFixedArray;

/**
 * Class to handle strings inside of XLSX files which are put to a specific shared strings file.
 *
 * @author Aspera GmbH
 */
class SharedStrings
{
    /**
     * Amount of array positions to add with each extension of the shared string cache array.
     * Larger values run into memory limits faster, lower values are a tiny bit worse on performance
     *
     * @var int
     */
    const SHARED_STRING_CACHE_ARRAY_SIZE_STEP = 100;

    /** @var string Filename (without path) of the shared strings XML file. */
    private $shared_strings_filename;

    /** @var string Path to the directory containing all shared strings files used by this instance. Includes trailing slash. */
    private $shared_strings_directory;

    /** @var OoxmlReader XML reader object for the shared strings XML file */
    private $shared_strings_reader;

    /** @var SharedStringsConfiguration Configuration of shared string reading and caching behaviour. */
    private $shared_strings_configuration;

    /** @var SplFixedArray Shared strings cache, if the number of shared strings is low enough */
    private $shared_string_cache;

    /**
     * Array of SharedStringsOptimizedFile instances containing filenames and associated data for shared strings that
     * were not saved to $shared_string_cache. Files contain values in seek-optimized format. (one entry per line, JSON encoded)
     * Key per element: the index of the first string contained within the file.
     *
     * @var array
     */
    private $prepared_shared_string_files = array();

    /** @var int The total number of shared strings available in the file. */
    private $shared_string_count = 0;

    /** @var int The shared string index the shared string reader is currently pointing at. */
    private $shared_string_index = 0;

    /** @var string|null Temporary cache for the last value that was read from the shared strings xml file. */
    private $last_shared_string_value;

    /**
     * SharedStrings constructor. Prepares the data stored within the given shared string file for reading.
     *
     * @param   string                      $shared_strings_directory       Directory of the shared strings file.
     * @param   string                      $shared_strings_filename        Filename of the shared strings file.
     * @param   SharedStringsConfiguration  $shared_strings_configuration   Configuration for shared string reading and
     *                                                                      caching behaviour.
     *
     * @throws  RuntimeException
     */
    public function __construct(
        $shared_strings_directory,
        $shared_strings_filename,
        SharedStringsConfiguration $shared_strings_configuration = null
    ) {
        $this->shared_strings_configuration = $shared_strings_configuration ?: new SharedStringsConfiguration();
        $this->shared_strings_directory = $shared_strings_directory;
        $this->shared_strings_filename = $shared_strings_filename;
        if (is_readable($this->shared_strings_directory . $this->shared_strings_filename)) {
            $this->prepareSharedStrings();
        }
    }

    /**
     * Closes all file pointers managed by this SharedStrings instance.
     * Note: Does not unlink temporary files. Use getTempFiles() to retrieve the list of created temp files.
     */
    public function close()
    {
        if ($this->shared_strings_reader && $this->shared_strings_reader instanceof OoxmlReader) {
            $this->shared_strings_reader->close();
            $this->shared_strings_reader = null;
        }
        /** @var SharedStringsOptimizedFile $file_data */
        foreach ($this->prepared_shared_string_files as $file_data) {
            $file_data->closeHandle();
        }

        $this->shared_strings_directory = null;
        $this->shared_strings_filename = null;
    }

    /**
     * @param SharedStringsConfiguration $configuration
     */
    public function setSharedStringsConfiguration(SharedStringsConfiguration $configuration)
    {
        $this->shared_strings_configuration = $configuration;
    }

    /**
     * Returns a list of all temporary work files created in this SharedStrings instance.
     *
     * @return array List of temporary files; With absolute paths.
     */
    public function getTempFiles()
    {
        $ret = array();
        /** @var SharedStringsOptimizedFile $file_details */
        foreach ($this->prepared_shared_string_files as $file_details) {
            $ret[] = $file_details->getFile();
        }
        return $ret;
    }

    /**
     * Retrieves a shared string value by its index
     *
     * @param   int     $target_index   Shared string index
     * @return  string  Shared string of the given index
     *
     * @throws  RuntimeException
     */
    public function getSharedString($target_index)
    {
        // If index of the desired string is larger than possible, don't even bother.
        if ($this->shared_string_count && ($target_index >= $this->shared_string_count)) {
            return '';
        }

        // Read from RAM cache?
        if ($this->shared_strings_configuration->getUseCache() && isset($this->shared_string_cache[$target_index])) {
            return $this->shared_string_cache[$target_index];
        }

        // Read from optimized files?
        if ($this->shared_strings_configuration->getUseOptimizedFiles()) {
            $result = $this->getStringFromOptimizedFile($target_index);
            if ($result !== null) {
                return $result;
            }
        }

        // No cache and no optimized files; Read directly from original XML
        return $this->getStringFromOriginalSharedStringFile($target_index);
    }

    /**
     * Attempts to retrieve a string from the optimized shared string files.
     * May return null if unsuccessful.
     *
     * @param   int $target_index
     * @return  null|string
     *
     * @throws  RuntimeException
     */
    private function getStringFromOptimizedFile($target_index)
    {
        // Determine the target file to read from, given the smallest index obtainable from it.
        $index_of_target_file = null;
        foreach (array_keys($this->prepared_shared_string_files) as $lowest_index) {
            if ($lowest_index > $target_index) {
                break; // Because the array is ksorted, we can assume that we've found our value at this point.
            }
            $index_of_target_file = $lowest_index;
        }
        if ($index_of_target_file === null) {
            return null;
        }

        /** @var SharedStringsOptimizedFile $file_data */
        $file_data = $this->prepared_shared_string_files[$index_of_target_file];

        // Determine our target line in the target file
        $target_index_in_file = $target_index - $index_of_target_file; // note: $index_of_target_file is also the index of the first string within the file
        if ($file_data->getHandleCurrentIndex() == $target_index_in_file) {
            // tiny optimization; If a previous seek already evaluated the target value, return it immediately
            return $file_data->getValueAtCurrentIndex();
        }

        // We found our target file to read from. Open a file handle or retrieve an already opened one.
        $target_handle = $file_data->getHandle();
        if (!$target_handle) {
            $target_handle = $file_data->openHandle('rb');
        }

        // Potentially rewind the file handle.
        if ($file_data->getHandleCurrentIndex() > $target_index_in_file) {
            // Our file pointer points at an index after the one we're looking for; Rewind the file pointer
            $target_handle = $file_data->rewindHandle();
        }

        // Walk through the file up to the index we're looking for and return its value
        $file_line = null;
        while ($file_data->getHandleCurrentIndex() < $target_index_in_file) {
            $file_data->increaseHandleCurrentIndex();
            $file_line = fgets($target_handle);
            if ($file_line === false) {
                break; // unexpected EOF; Silent fallback to original shared string file.
            }
        }
        if (is_string($file_line) && $file_line !== '') {
            $file_line = json_decode($file_line);

            if ($this->shared_strings_configuration->getKeepFileHandles()) {
                $file_data->setValueAtCurrentIndex($file_line);
            } else {
                $file_data->closeHandle();
            }

            return $file_line;
        }

        return null;
    }

    /**
     * Retrieves a shared string from the original shared strings XML file.
     *
     * @param   int $target_index
     * @return  null|string
     */
    private function getStringFromOriginalSharedStringFile($target_index)
    {
        // If the desired index equals the current, return cached result.
        if ($target_index === $this->shared_string_index && $this->last_shared_string_value !== null) {
            return $this->last_shared_string_value;
        }

        // If the desired index is before the current, rewind the XML.
        if ($this->shared_strings_reader && $this->shared_string_index > $target_index) {
            $this->shared_strings_reader->close();
            $this->shared_strings_reader = null;
        }

        // Initialize reader, if not already initialized.
        if (!$this->shared_strings_reader) {
            $this->initSharedStringsReader();
        }

        // Move reader to the next <si> node, if it isn't already pointing at one.
        if (!$this->shared_strings_reader->matchesElement('si') || $this->shared_strings_reader->isClosingTag()) {
            $found_next_si_node = false;
            while ($this->shared_strings_reader->read()) {
                if ($this->shared_strings_reader->matchesElement('si') && !$this->shared_strings_reader->isClosingTag()) {
                    $found_next_si_node = true;
                    break;
                }
            }
            if (!$found_next_si_node) {
                // Unexpected EOF; The given sharedString index could not be found.
                $this->shared_strings_reader->close();
                $this->shared_strings_reader = null;
                return '';
            }
            $this->shared_string_index++;
        }

        // Move to the <si> node with the desired index
        $eof_reached = false;
        while (!$eof_reached && $this->shared_string_index < $target_index) {
            $eof_reached = !$this->shared_strings_reader->nextNsId('si');
            $this->shared_string_index++;
        }
        if ($eof_reached) {
            // Unexpected EOF; The given sharedString index could not be found.
            $this->shared_strings_reader->close();
            $this->shared_strings_reader = null;
            return '';
        }

        // Extract the value from the shared string
        $matched_elements = array(
            't'  => array('t'),
            'si' => array('si')
        );
        $value = '';
        while ($this->shared_strings_reader->read()) {
            switch ($this->shared_strings_reader->matchesOneOfList($matched_elements)) {
                // <t> - Read the shared string value contained within the element.
                case 't':
                    if ($this->shared_strings_reader->isClosingTag()) {
                        continue 2;
                    }
                    $value .= $this->shared_strings_reader->readString();
                    break;

                // </si> - End of entry. Abort further reading.
                case 'si':
                    if ($this->shared_strings_reader->isClosingTag()) {
                        break 2;
                    }
                    break;
            }
        }

        if (!$this->shared_strings_configuration->getKeepFileHandles()) {
            $this->shared_strings_reader->close();
            $this->shared_strings_reader = null;
        }

        $this->last_shared_string_value = $value;
        return $value;
    }

    /**
     * Initializes the shared strings XML reader object with the proper settings.
     * Also initializes all related tracking properties.
     */
    private function initSharedStringsReader()
    {
        $this->shared_strings_reader = new OoxmlReader();
        $this->shared_strings_reader->setDefaultNamespaceIdentifierElements(OoxmlReader::NS_XLSX_MAIN);
        $this->shared_strings_reader->setDefaultNamespaceIdentifierAttributes(OoxmlReader::NS_NONE);
        $this->shared_strings_reader->open($this->shared_strings_directory . $this->shared_strings_filename);
        $this->shared_string_index = -1;
        $this->last_shared_string_value = null;
    }

    /**
     * Perform optimizations to increase performance of shared string determination operations.
     * Loads shared string data into RAM up to the configured memory limit. Stores additional shared string data
     * in seek-optimized additional files on the filesystem in order to lower seek times.
     *
     * @return void
     *
     * @throws RuntimeException
     */
    private function prepareSharedStrings()
    {
        $this->initSharedStringsReader();

        // Obtain number of shared strings available
        while ($this->shared_strings_reader->read()) {
            if ($this->shared_strings_reader->matchesElement('sst')) {
                $this->shared_string_count = $this->shared_strings_reader->getAttributeNsId('uniqueCount');
                break;
            }
        }
        if (!$this->shared_string_count) {
            // No shared strings available, no preparation necessary
            $this->shared_strings_reader->close();
            $this->shared_strings_reader = null;
            return;
        }

        if ($this->shared_strings_configuration->getUseCache()) {
            // This is why we ask for at least 8 KB of memory. Lower values may already exceed the limit with this assignment:
            $this->shared_string_cache = new SplFixedArray(self::SHARED_STRING_CACHE_ARRAY_SIZE_STEP);
        }

        // Prepare working through the XML file. Declare as many variables as makes sense upfront, for more accurate memory usage retrieval.
        $string_index = 0;
        $string_value = '';
        $write_to_cache = $this->shared_strings_configuration->getUseCache();
        $cache_max_size_byte = $this->shared_strings_configuration->getCacheSizeKilobyte() * 1024;
        $matched_elements = array(
            'si' => array('si'),
            't' => array('t')
        );

        $start_memory_byte = memory_get_usage(false); // Note: Get current memory usage as late as possible. Read: Now.

        // Work through the XML file and cache/reformat/move string data, according to configuration and situation
        while ($this->shared_strings_reader->read()) {
            switch ($this->shared_strings_reader->matchesOneOfList($matched_elements)) {
                // <t> - Read shared string value portion contained within the element.
                case 't':
                    if (!$this->shared_strings_reader->isClosingTag()) {
                        $string_value .= $this->shared_strings_reader->readString();
                    }
                    break;

                // </si> - Write previously read string value to cache.
                case 'si':
                    if (!$this->shared_strings_reader->isClosingTag()) {
                        break;
                    }
                    if ($write_to_cache) {
                        $cache_current_memory_byte = memory_get_usage(false) - $start_memory_byte;
                        if ($cache_current_memory_byte > $cache_max_size_byte) {
                            // transition from "cache everything" to "memory exhausted, stop caching":
                            $this->shared_string_cache->setSize($string_index); // finalize array size
                            $write_to_cache = false;
                        }
                    }
                    $this->prepareSingleSharedString($string_index, $string_value, $write_to_cache);
                    $string_index++;
                    $string_value = '';
                    break;
            }
        }

        // Small optimization: Sort shared string files by lowest included key for slightly faster reading.
        ksort($this->prepared_shared_string_files);

        // Close all no longer necessary file handles
        $this->shared_strings_reader->close();
        $this->shared_strings_reader = null;

        /** @var SharedStringsOptimizedFile $file_data */
        foreach ($this->prepared_shared_string_files as $file_data) {
            $file_data->closeHandle();
        }
    }

    /**
     * Stores the given shared string either in internal cache or in a seek optimized file, depending on the
     * current configuration and status of the internal cache.
     *
     * @param   int     $index
     * @param   string  $string
     * @param   bool    $write_to_cache
     *
     * @throws  RuntimeException
     */
    private function prepareSingleSharedString($index, $string, $write_to_cache = false)
    {
        if ($write_to_cache) {
            // Caching enabled and there's still memory available; Write to internal cache.
            if ($index + 1 > $this->shared_string_cache->getSize()) {
                $this->shared_string_cache->setSize($this->shared_string_cache->getSize() + self::SHARED_STRING_CACHE_ARRAY_SIZE_STEP);
            }
            $this->shared_string_cache[$index] = $string;
            return;
        }

        if (!$this->shared_strings_configuration->getUseOptimizedFiles()) {
            // No preparation possible. This value will have to be read from the original shared string XML file.
            return;
        }

        // Caching not possible. Write shared string to seek-optimized file instead.

        // Check if we have an already existing file that still has room for more entries in it.
        /** @var SharedStringsOptimizedFile $newest_file_data */
        $newest_file_data = null;
        $newest_file_is_full = false;
        $shared_string_file_index = null;
        if (count($this->prepared_shared_string_files) > 0) {
            $shared_string_file_index = max(array_keys($this->prepared_shared_string_files));
            $newest_file_data = $this->prepared_shared_string_files[$shared_string_file_index];
            if ($newest_file_data->getCount() >= $this->shared_strings_configuration->getOptimizedFileEntryCount()) {
                $newest_file_is_full = true;
            }
        }

        $create_new_file = !$newest_file_data || $newest_file_is_full;
        if ($create_new_file) {
            // Assemble new filename; Add random hash to avoid conflicts for when the target directory is also used by other processes.
            $hash = base_convert(mt_rand(36 ** 4, (36 ** 5) - 1), 10, 36); // Possible results: "10000" - "zzzzz"
            $newest_file_data = new SharedStringsOptimizedFile();
            $filename_without_suffix = preg_replace('~(.+)\.[^./]$~', '$1', $this->shared_strings_filename);
            $newest_file_data->setFile($this->shared_strings_directory . $filename_without_suffix . '_tmp_' . $index . '_' . $hash . '.txt');
            $fhandle = $newest_file_data->openHandle('wb');
            $this->prepared_shared_string_files[$index] = $newest_file_data;
        } else {
            // Append shared string to the newest file.
            $fhandle = $newest_file_data->getHandle();
            if (!$fhandle) {
                $fhandle = $newest_file_data->openHandle('ab');
            }
        }

        // Write shared string to the chosen file
        if (fwrite($fhandle, json_encode($string) . PHP_EOL) === false) {
            throw new RuntimeException('Could not write shared string to temporary file.');
        }
        $newest_file_data->increaseCount();

        if (!$this->shared_strings_configuration->getKeepFileHandles()) {
            $newest_file_data->closeHandle();
        }
    }
}


#namespace Aspera\Spreadsheet\XLSX;

#use InvalidArgumentException;

/**
 * Holds all configuration options related to shared string related behaviour
 *
 * @author Aspera GmbH
 */
class SharedStringsConfiguration
{
    /**
     * If true: Allow caching shared strings to RAM to increase performance.
     *
     * @var bool
     */
    private $use_cache = true;

    /**
     * Maximum allowed RAM consumption for shared string cache, in kilobyte. (Minimum: 8 KB)
     * Once exceeded, additional shared strings will not be written to RAM and instead get read from file as needed.
     * Note that this is a "soft" limit that only applies to the main cache. The application may slightly exceed it.
     *
     * @var int
     */
    private $cache_size_kilobyte = 256;

    /**
     * If true: Allow creation of new files to reduce seek times for non-cached shared strings.
     *
     * @var bool
     */
    private $use_optimized_files = true;

    /**
     * Amount of shared strings to store per seek optimized shared strings file.
     * Lower values result in higher performance at the cost of more temporary files being created.
     * At extremely low values (< 10) you might be better off increasing the cache size.
     *
     * @var int
     */
    private $optimized_file_entry_count = 2500;

    /**
     * If true: file pointers to shared string files are kept open for more efficient reads.
     * Causes higher memory consumption, especially if $optimized_file_entry_count is low.
     *
     * @var bool
     */
    private $keep_file_handles = true;

    /**
     * Enable/disable caching of shared string values in RAM.
     *
     * @param   bool    $new_use_cache_value
     *
     * @throws  InvalidArgumentException
     */
    public function setUseCache($new_use_cache_value)
    {
        if (!is_bool($new_use_cache_value)) {
            throw new InvalidArgumentException('Invalid parameter value; Expected a boolean.');
        }
        $this->use_cache = $new_use_cache_value;
    }


    /**
     * Set the maximum size of the internal shared string cache, in kilobyte. (minimum: 8 KB)
     * Note that this is a soft limit; Depending on circumstances, it might be exceeded by a few byte/kilobyte.
     *
     * @param   int $new_max_size
     *
     * @throws  InvalidArgumentException
     */
    public function setCacheSizeKilobyte($new_max_size)
    {
        if (!is_numeric($new_max_size) || $new_max_size < 8) {
            throw new InvalidArgumentException('Invalid parameter value; Expected a positive number equal to or greater than 8.');
        }
        $this->cache_size_kilobyte = (int)$new_max_size;
    }

    /**
     * Enable/disable the creation of new temporary files for the purpose of optimizing shared string seek performance.
     *
     * @param   bool    $new_use_files_value
     *
     * @throws  InvalidArgumentException
     */
    public function setUseOptimizedFiles($new_use_files_value)
    {
        if (!is_bool($new_use_files_value)) {
            throw new InvalidArgumentException('Invalid parameter value; Expected a boolean.');
        }
        $this->use_optimized_files = $new_use_files_value;
    }

    /**
     * Set the amount of entries to be stored per single optimized shared string file.
     * Adjusting this value has no effect if the creation of optimized shared string files is disabled.
     *
     * @param   int $new_entry_count
     *
     * @throws  InvalidArgumentException
     */
    public function setOptimizedFileEntryCount($new_entry_count)
    {
        if (!is_numeric($new_entry_count) || $new_entry_count <= 0) {
            throw new InvalidArgumentException('Invalid parameter value; Expected a positive number.');
        }
        $this->optimized_file_entry_count = $new_entry_count;
    }

    /**
     * Enable/disable keeping file pointers to shared string files open to achieve more efficient file reads.
     *
     * @param   bool    $new_keep_file_pointers_value
     *
     * @throws  InvalidArgumentException
     */
    public function setKeepFileHandles($new_keep_file_pointers_value)
    {
        if (!is_bool($new_keep_file_pointers_value)) {
            throw new InvalidArgumentException('Invalid parameter value; Expected a boolean.');
        }
        $this->keep_file_handles = $new_keep_file_pointers_value;
    }

    /**
     * @return bool
     */
    public function getUseCache()
    {
        return $this->use_cache;
    }

    /**
     * @return int
     */
    public function getCacheSizeKilobyte()
    {
        return $this->cache_size_kilobyte;
    }

    /**
     * @return bool
     */
    public function getUseOptimizedFiles()
    {
        return $this->use_optimized_files;
    }

    /**
     * @return int
     */
    public function getOptimizedFileEntryCount()
    {
        return $this->optimized_file_entry_count;
    }

    /**
     * @return bool
     */
    public function getKeepFileHandles()
    {
        return $this->keep_file_handles;
    }
}


#namespace Aspera\Spreadsheet\XLSX;

#use RuntimeException;

/**
 * Data object to hold all data corresponding to a single optimized shared string file (not the original XML file).
 *
 * @author Aspera GmbH
 */
class SharedStringsOptimizedFile
{
    /** @var string Complete path to the file. */
    private $file = '';

    /** @var resource File handle to access the file contents with. */
    private $handle;

    /** @var int Index of the line the handle currently points at. (Only used during reading from the file) */
    private $handle_current_index = -1;

    /** @var string The shared string value corresponding to the current index. (Only used during reading from the file) */
    private $value_at_current_index = '';

    /** @var int Total number of shared strings contained within the file. */
    private $count = 0;

    /**
     * @return string
     */
    public function getFile()
    {
        return $this->file;
    }

    /**
     * @param string $file
     */
    public function setFile($file)
    {
        $this->file = $file;
    }

    /**
     * @return resource
     */
    public function getHandle()
    {
        return $this->handle;
    }

    /**
     * @param resource $handle
     */
    public function setHandle($handle)
    {
        $this->handle = $handle;
    }

    /**
     * @return int
     */
    public function getHandleCurrentIndex()
    {
        return $this->handle_current_index;
    }

    /**
     * @param int $handle_current_index
     */
    public function setHandleCurrentIndex($handle_current_index)
    {
        $this->handle_current_index = $handle_current_index;
    }

    /**
     * Increase current index of handle by 1.
     */
    public function increaseHandleCurrentIndex()
    {
        $this->handle_current_index++;
    }

    /**
     * @return string
     */
    public function getValueAtCurrentIndex()
    {
        return $this->value_at_current_index;
    }

    /**
     * @param string $value_at_current_index
     */
    public function setValueAtCurrentIndex($value_at_current_index)
    {
        $this->value_at_current_index = $value_at_current_index;
    }

    /**
     * @return int
     */
    public function getCount()
    {
        return $this->count;
    }

    /**
     * @param int $count
     */
    public function setCount($count)
    {
        $this->count = $count;
    }

    /**
     * Increase count of elements contained within the file by 1.
     */
    public function increaseCount()
    {
        $this->count++;
    }

    /**
     * Opens a file handle to the file with the given file access mode.
     * If a file handle is currently still open, closes it first.
     *
     * @param   string      $mode
     * @return  resource    The newly opened file handle
     *
     * @throws  RuntimeException
     */
    public function openHandle($mode)
    {
        $this->closeHandle();
        $new_handle = @fopen($this->getFile(), $mode);
        if (!$new_handle) {
            throw new RuntimeException(
                'Could not open file handle for optimized shared string file with mode ' . $mode . '.'
            );
        }
        $this->setHandle($new_handle);
        return $this->getHandle();
    }

    /**
     * Properly closes the current file handle, if it is currently opened.
     */
    public function closeHandle()
    {
        if (!$this->handle) {
            return; // Nothing to close
        }
        fclose($this->handle);
        $this->handle = null;
        $this->handle_current_index = -1;
        $this->value_at_current_index = '';
    }

    /**
     * Properly rewinds the current file handle and all associated internal data.
     *
     * @return  resource    The rewound file handle
     *
     * @throws  RuntimeException
     */
    public function rewindHandle()
    {
        if (!$this->handle) {
            throw new RuntimeException('Could not rewind file handle; There is no file handle currently open.');
        }
        rewind($this->handle);
        $this->handle_current_index = -1;
        $this->value_at_current_index = null;
        return $this->handle;
    }
}


#namespace Aspera\Spreadsheet\XLSX;

/**
 * Data object for worksheet related data
 *
 * @author Aspera GmbH
 */
class Worksheet
{
    /** @var string */
    private $name;

    /** @var string Relationship ID of this worksheet for matching with workbook data. */
    private $relationship_id;

    /**
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }
    
    public function fputcsv($a,$b)
    {
        return fputcsv($a,$b);
    }

    /**
     * @param string $name
     */
    public function setName($name)
    {
        $this->name = $name;
    }

    /**
     * @return string
     */
    public function getRelationshipId()
    {
        return $this->relationship_id;
    }

    /**
     * @param string $relationship_id
     */
    public function setRelationshipId($relationship_id)
    {
        $this->relationship_id = $relationship_id;
    }
}


#namespace Aspera\Spreadsheet\XLSX;

#use ZipArchive;

/**
 * Data object containing all data related to a single 1:1 relationship declaration
 *
 * @author Aspera GmbH
 */
class RelationshipElement
{
    /** @var string Internal identifier of this file part */
    private $id;

    /** @var bool Element validity flag; If false, this element was not found or might be corrupted. */
    private $is_valid;

    /** @var string Path to this element, as per the context its information was retrieved from. */
    private $original_path;

    /** @var string Absolute path to the file associated with this element for access. */
    private $access_path;

    /**
     * @return string
     */
    public function getId()
    {
        return $this->id;
    }

    /**
     * @param string $id
     */
    public function setId($id)
    {
        $this->id = $id;
    }

    /**
     * @return bool
     */
    public function isValid()
    {
        return $this->is_valid;
    }

    /**
     * @param bool $is_valid
     */
    public function setIsValid($is_valid)
    {
        $this->is_valid = $is_valid;
    }

    /**
     * @return string
     */
    public function getOriginalPath()
    {
        return $this->original_path;
    }

    /**
     * @param string $original_path
     */
    public function setOriginalPath($original_path)
    {
        $this->original_path = $original_path;
    }

    /**
     * @return string
     */
    public function getAccessPath()
    {
        return $this->access_path;
    }

    /**
     * @param string $access_path
     */
    public function setAccessPath($access_path)
    {
        $this->access_path = $access_path;
    }

    /**
     * Checks the given zip file for the element described by this object and sets validity flag accordingly.
     *
     * @param ZipArchive $zip
     */
    public function setValidityViaZip($zip)
    {
        $this->setIsValid($zip->locateName($this->getOriginalPath()) !== false);
    }
}

#namespace Aspera\Spreadsheet\XLSX;

#use XMLReader;
#use InvalidArgumentException;

/**
 * Extension of XMLReader to ease parsing of XML files of the OOXML specification.
 *
 * Depending on edition, namespaceUris in OOXML documents can be entirely different. Besides adding extra
 * matching overhead, this makes custom-made documents that are employing their own namespace rules a bit
 * complicated to read correctly. To mitigate the impact of that, this wrapper of XMLReader supplies methods
 * that deal with these issues automatically.
 *
 * @author Aspera GmbH
 */
class OoxmlReader extends XMLReader
{
    /**
     * Identifiers of supported OOXML Namespaces.
     * Use these instead of namespaceUris when checking for elements that are part of OOXML namespaces.
     *
     * @var array NS_NONE Also known as the "empty" namespace. All attributes always default to this.
     * @var array XMLNS_XLSX_MAIN Root namespace of most XLSX documents.
     * @var array XMLNS_RELATIONSHIPS_DOCUMENTLEVEL Namespace used for references to relationship documents.
     * @var array XMLNS_RELATIONSHIPS_PACKAGELEVEL Root namespace used within relationship documents.
     */
    const NS_NONE = '';
    const NS_XLSX_MAIN = 'xlsx_main';
    const NS_RELATIONSHIPS_DOCUMENTLEVEL = 'relationships_documentlevel';
    const NS_RELATIONSHIPS_PACKAGELEVEL = 'relationships_packagelevel';

    /** @var array Format: $namespace_list[-XMLNS_IDENTIFIER-][-INTRODUCING_EDITION_OF_SPECIFICATION-] = -NAMESPACE_URI- */
    private $namespace_list;

    /** @var string One of the NS_ constants that will be used if methods requiring a NsId for an element tag do not get one delivered. */
    private $default_namespace_identifier_elements;

    /** @var string One of the NS_ constants that will be used if methods requiring a NsId for an attribute do not get one delivered. */
    private $default_namespace_identifier_attributes;

    /**
     * Initialize $this->namespace_list.
     */
    private function initNamespaceList()
    {
        $this->namespace_list = array(
            self::NS_NONE => array(''),
            self::NS_XLSX_MAIN => array(
                1 => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                3 => 'http://purl.oclc.org/ooxml/spreadsheetml/main'
            ),
            self::NS_RELATIONSHIPS_DOCUMENTLEVEL => array(
                1 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                3 => 'http://purl.oclc.org/ooxml/officeDocument/relationships'
            ),
            self::NS_RELATIONSHIPS_PACKAGELEVEL  => array(
                1 => 'http://schemas.openxmlformats.org/package/2006/relationships',
                3 => 'http://purl.oclc.org/ooxml/officeDocument/relationships' // Note: Same as DOCUMENTLEVEL
            )
        );
    }

    /**
     * Sets the default namespace_identifier for element tags,
     * to be used when methods requiring a namespace_identifier are not given one.
     *
     * @param   string  $namespace_identifier
     *
     * @throws  InvalidArgumentException
     */
    public function setDefaultNamespaceIdentifierElements($namespace_identifier)
    {
        if (!isset($this->namespace_list[$namespace_identifier])) {
            throw new InvalidArgumentException('unknown namespace identifier [' . $namespace_identifier . ']');
        }
        $this->default_namespace_identifier_elements = $namespace_identifier;
    }

    /**
     * Sets the default namespace_identifier for element attributes,
     * to be used when methods requiring a namespace_identifier are not given one.
     *
     * @param   string  $namespace_identifier
     *
     * @throws  InvalidArgumentException
     */
    public function setDefaultNamespaceIdentifierAttributes($namespace_identifier)
    {
        if (!isset($this->namespace_list[$namespace_identifier])) {
            throw new InvalidArgumentException('unknown namespace identifier [' . $namespace_identifier . ']');
        }
        $this->default_namespace_identifier_attributes = $namespace_identifier;
    }

    public function __construct()
    {
        $this->initNamespaceList();
        // Note: No parent::__construct() - XMLReader does not have its own constructor.
    }

    /**
     * Checks if the element the reader is currently pointed at is of the given local_name with a namespace_uri
     * that matches the list of namespaces identified by the given namespace_identifier constant.
     *
     * @param   string      $local_name
     * @param   string|null $namespace_identifier   NULL = Fallback to $this->default_namespace_identifier_elements
     * @return  bool
     *
     * @throws  InvalidArgumentException
     */
    public function matchesElement($local_name, $namespace_identifier = null)
    {
        return $this->localName === $local_name
            && $this->matchesNamespace($namespace_identifier);
    }

    /**
     * Checks if any of the given list of elements is matched by the current element.
     * Returns the array key of the element that matched, or false if none matched.
     *
     * @param   array   $list_of_elements   Format: array([MATCH_1_ID] => array(LOCAL_NAME_1, NAMESPACE_ID_1), ...)
     * @return  mixed|false If no match was found: false. Otherwise, the parameter array's key of the element definition that matched.
     */
    public function matchesOneOfList($list_of_elements)
    {
        foreach ($list_of_elements as $one_element_key => $one_element) {
            $parameter_count = count($one_element);
            if ($parameter_count < 1 || $parameter_count > 2) {
                throw new InvalidArgumentException('Invalid definition of element. Expected 1 or 2 parameters, got [' . $parameter_count . '].');
            }
            if ($this->localName !== $one_element[0]) {
                continue;
            }
            if (!isset($one_element[1])) {
                $one_element[1] = null; // default $namespace_identifier value
            }
            if ($this->matchesNamespace($one_element[1])) {
                return $one_element_key;
            }
        }
        return false;
    }

    /**
     * Checks if the element the reader is currently pointed at contains an element with a namespace_uri
     * that matches the list of namespaces identified by the given namespace_identifier constant.
     *
     * @param   string|null $namespace_identifier   NULL = Fallback to $this->default_namespace_identifier_elements
     * @param   bool        $for_attribute          Determines the scope of validation; true: attribute, false: element tag
     * @return  bool
     *
     * @throws  InvalidArgumentException
     */
    public function matchesNamespace($namespace_identifier = null, $for_attribute = false)
    {
        return in_array(
            $this->namespaceURI,
            $this->namespace_list[$this->validateNamespaceIdentifier($namespace_identifier, $for_attribute)],
            true
        );
    }

    /**
     * Checks if the current element is a closing tag / END_ELEMENT.
     *
     * @return bool
     */
    public function isClosingTag() {
        return $this->nodeType === OoxmlReader::END_ELEMENT;
    }

    /**
     * Extension of getAttributeNs that checks with a namespace_identifier rather than a specific namespace_uri.
     *
     * @param   string      $local_name
     * @param   string|null $namespace_identifier   NULL = Fallback to $this->default_namespace_identifier_elements
     * @return  NULL|string
     *
     * @throws  InvalidArgumentException
     */
    public function getAttributeNsId($local_name, $namespace_identifier = null)
    {
        $namespace_identifier = $this->validateNamespaceIdentifier($namespace_identifier, true);

        $ret_value = null;
        foreach ($this->namespace_list[$namespace_identifier] as $namespace_uri) {
            $moved_successfully = ($namespace_uri === '')
                ? $this->moveToAttribute($local_name)
                : $this->moveToAttributeNs($local_name, $namespace_uri);
            if ($moved_successfully) {
                $ret_value = $this->value;
                break;
            }
        }
        $this->moveToElement();

        return $ret_value;
    }

    /**
     * Moves to the next node matching the given criteria.
     *
     * @param   string      $local_name
     * @param   string|null $namespace_identifier
     * @return  bool
     */
    public function nextNsId($local_name, $namespace_identifier = null)
    {
        while ($this->next($local_name)) {
            if ($this->matchesNamespace($namespace_identifier)) {
                return true;
            }
        }

        return false;
    }

    /**
     * Checks if the given namespace_identifier is valid. If null is given, will try to fallback to
     * $this->default_namespace_identifier_elements. Returns the correct namespace_identifier for further usage.
     *
     * @param   string|null $namespace_identifier
     * @param   bool        $for_attribute          Determines the default namespace_identifier to fallback to.
     * @return  string
     *
     * @throws  InvalidArgumentException
     */
    private function validateNamespaceIdentifier($namespace_identifier, $for_attribute = false)
    {
        if ($namespace_identifier === null) {
            $default_namespace_identifier = ($for_attribute)
                ? $this->default_namespace_identifier_attributes
                : $this->default_namespace_identifier_elements;
            if ($default_namespace_identifier === null) {
                throw new InvalidArgumentException('no namespace identifier given');
            }

            return $default_namespace_identifier;
        } elseif (!isset($this->namespace_list[$namespace_identifier])) {
            throw new InvalidArgumentException('unknown namespace identifier [' . $namespace_identifier . ']');
        }

        return $namespace_identifier;
    }
}


#namespace Aspera\Spreadsheet\XLSX;

#use RuntimeException;
#use ZipArchive;

/**
 * Functionality to work with relationship data (.rels files)
 * Also contains all relationship data it previously evaluated for structured retrieval.
 *
 * @author Aspera GmbH
 */
class RelationshipData
{
    /**
     * Directory separator character used in zip file internal paths.
     * Is supposed to always be a forward slash, even on systems with a different directory separator (e.g. Windows).
     *
     * @var string ZIP_DIR_SEP
     */
    const ZIP_DIR_SEP = '/';

    /** @var RelationshipElement Workbook file meta information. Only one element exists per file. */
    private $workbook;

    /** @var array Worksheet files meta information, saved as a list of RelationshipElement instances. */
    private $worksheets = array();

    /** @var array SharedStrings files meta information, saved as a list of RelationshipElement instances. */
    private $shared_strings = array();

    /** @var array Styles files meta information, saved as a list of RelationshipElement instances. */
    private $styles = array();

    /**
     * Returns the workbook relationship element, if a valid one has been obtained previously.
     * Returns null otherwise.
     *
     * @return null|RelationshipElement
     */
    public function getWorkbook()
    {
        if (isset($this->workbook) && $this->workbook->isValid()) {
            return $this->workbook;
        }
        return null;
    }

    /**
     * Returns data of all found valid shared string elements.
     * Returns array of RelationshipElement elements.
     *
     * @return array[RelationshipElement]
     */
    public function getSharedStrings()
    {
        $return_list = array();
        foreach ($this->shared_strings as $shared_string_element) {
            if ($shared_string_element->isValid()) {
                $return_list[] = $shared_string_element;
            }
        }
        return $return_list;
    }

    /**
     * Returns all worksheet data of all found valid worksheet elements.
     * Returns array of RelationshipElement elements.
     *
     * @return array
     */
    public function getWorksheets()
    {
        $return_list = array();
        foreach ($this->worksheets as $worksheet_element) {
            if ($worksheet_element->isValid()) {
                $return_list[] = $worksheet_element;
            }
        }
        return $return_list;
    }

    /**
     * Returns all styles data of all found valid styles elements.
     * Returns array of RelationshipElement elements
     *
     * @return array
     */
    public function getStyles()
    {
        $return_list = array();
        foreach ($this->styles as $styles_element) {
            if ($styles_element->isValid()) {
                $return_list[] = $styles_element;
            }
        }
        return $return_list;
    }

    /**
     * Navigates through the XLSX file using .rels files, gathering up found file parts along the way.
     * Results are saved in internal variables for later retrieval.
     *
     * @param   ZipArchive  $zip    Handle to zip file to read relationship data from
     *
     * @throws  RuntimeException
     */
    public function __construct(ZipArchive $zip)
    {
        // Start with root .rels file. It will point us towards the worksheet file.
        $root_rel_file = self::toRelsFilePath(''); // empty string returns root path
        $this->evaluateRelationshipFromZip($zip, $root_rel_file);

        // Quick check: Workbook should have been retrieved from root relationship file.
        if (!isset($this->workbook) || !$this->workbook->isValid()) {
            throw new RuntimeException('Could not locate workbook data.');
        }

        // The workbook .rels file should point us towards all other required files.
        $workbook_rels_file_path = self::toRelsFilePath($this->workbook->getOriginalPath());
        $this->evaluateRelationshipFromZip($zip, $workbook_rels_file_path);
    }

    /**
     * Read through the .rels data of the given .rels file from the given zip handle
     * and save all included file data to internal variables.
     *
     * @param   ZipArchive  $zip
     * @param   string      $file_zipname
     *
     * @throws  RuntimeException
     */
    private function evaluateRelationshipFromZip(ZipArchive $zip, $file_zipname)
    {
        if ($zip->locateName($file_zipname) === false) {
            throw new RuntimeException('Could not read relationship data. File [' . $file_zipname . '] could not be found.');
        }

        $rels_reader = new OoxmlReader();
        $rels_reader->setDefaultNamespaceIdentifierElements(OoxmlReader::NS_RELATIONSHIPS_PACKAGELEVEL);
        $rels_reader->setDefaultNamespaceIdentifierAttributes(OoxmlReader::NS_NONE);
        $rels_reader->xml($zip->getFromName($file_zipname));
        while ($rels_reader->read() !== false) {
            if (!$rels_reader->matchesElement('Relationship') || $rels_reader->isClosingTag()) {
                // This element is not important to us. Skip.
                continue;
            }

            // Only the last part of the relationship type definition matters to us.
            $rels_type = $rels_reader->getAttributeNsId('Type');
            if (!preg_match('~([^/]+)/?$~', $rels_type, $type_regexp_matches)) {
                throw new RuntimeException(
                    'Invalid type definition found: [' . $rels_type . ']'
                    . ' Relationship could not be evaluated.'
                );
            }

            // Adjust target path (making it absolute without leading slash) so that we can easily use it for zip checks later.
            $target_path = $rels_reader->getAttributeNsId('Target');
            if (strpos($target_path, self::ZIP_DIR_SEP) === 0) {
                // target_path is already absolute, but we need to remove the leading slash.
                $target_path = substr($target_path, 1);
            } elseif (preg_match('~(.*' . self::ZIP_DIR_SEP . ')_rels~', $file_zipname, $path_matches)) {
                // target_path is relative. Add path of this .rels file to target path
                $target_path = $path_matches[1] . $target_path;
            }

            // Assemble and store element data
            $element_data = new RelationshipElement();
            $element_data->setId($rels_reader->getAttributeNsId('Id'));
            $element_data->setOriginalPath($target_path);
            $element_data->setValidityViaZip($zip);
            switch ($type_regexp_matches[1]) {
                case 'officeDocument':
                    $this->workbook = $element_data;
                    break;
                case 'worksheet':
                    $this->worksheets[] = $element_data;
                    break;
                case 'sharedStrings':
                    $this->shared_strings[] = $element_data;
                    break;
                case 'styles':
                    $this->styles[] = $element_data;
                    break;
                default:
                    // nop
                    break;
            }
        }
    }

    /**
     * Returns the path to the .rels file for the given file path.
     * Example: xl/workbook.xml => xl/_rels/workbook.xml.rels
     *
     * @param   string  $file_path
     * @return  string
     */
    private static function toRelsFilePath($file_path)
    {
        // Normalize directory separator character
        $file_path = str_replace('\\', self::ZIP_DIR_SEP, $file_path);

        // Split path in 2 parts around last dir seperator: [path/to/file]/[file.xml]
        $last_slash_pos = strrpos($file_path, '/');
        if ($last_slash_pos === false) {
            // No final slash; file.xml => _rels/file.xml.rels
            // This also implicitly handles the root .rels file, always found under "_rels/.rels"
            $file_path = '_rels/' . $file_path . '.rels';
        } elseif ($last_slash_pos == strlen($file_path) - 1) {
            // Trailing slash; some/folder/ => some/_rels/folder.rels
            $file_path = $file_path . '_rels/.rels';
        } else {
            // File with path; some/folder/file.xml => some/folder/_rels/file.xml.rels
            $file_path = preg_replace('~([^/]+)$~', '_rels/$1.rels', $file_path);
        }
        return $file_path;
    }
}


#namespace Aspera\Spreadsheet\XLSX;

#use Iterator;
#use Countable;
#use RuntimeException;
#use ZipArchive;
#use DateTime;
#use DateTimeZone;
#use DateInterval;
#use Exception;
#use InvalidArgumentException;

/**
 * Class for parsing XLSX files.
 *
 * @author Aspera GmbH
 * @author Martins Pilsetnieks
 */
class Reader implements Iterator, Countable
{
    /** @var array Base formats for XLSX documents, to be made available without former declaration. */
    const BUILTIN_FORMATS = array(
        0 => '',
        1 => '0',
        2 => '0.00',
        3 => '#,##0',
        4 => '#,##0.00',

        9  => '0%',
        10 => '0.00%',
        11 => '0.00E+00',
        12 => '# ?/?',
        13 => '# ??/??',
        14 => 'mm-dd-yy',
        15 => 'd-mmm-yy',
        16 => 'd-mmm',
        17 => 'mmm-yy',
        18 => 'h:mm AM/PM',
        19 => 'h:mm:ss AM/PM',
        20 => 'h:mm',
        21 => 'h:mm:ss',
        22 => 'm/d/yy h:mm',

        37 => '#,##0 ;(#,##0)',
        38 => '#,##0 ;[Red](#,##0)',
        39 => '#,##0.00;(#,##0.00)',
        40 => '#,##0.00;[Red](#,##0.00)',

        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mmss.0',
        48 => '##0.0E+0',
        49 => '@',

        // CHT & CHS
        27 => '[$-404]e/m/d',
        30 => 'm/d/yy',
        36 => '[$-404]e/m/d',
        50 => '[$-404]e/m/d',
        57 => '[$-404]e/m/d',

        // THA
        59 => 't0',
        60 => 't0.00',
        61 => 't#,##0',
        62 => 't#,##0.00',
        67 => 't0%',
        68 => 't0.00%',
        69 => 't# ?/?',
        70 => 't# ??/??'
    );

    /** @var array Conversion matrix to convert XLSX date formats to PHP date formats. */
    const DATE_REPLACEMENTS = array(
        'All' => array(
            '\\'    => '',
            'am/pm' => 'A',
            'yyyy'  => 'Y',
            'yy'    => 'y',
            'mmmmm' => 'M',
            'mmmm'  => 'F',
            'mmm'   => 'M',
            ':mm'   => ':i',
            'mm'    => 'm',
            'm'     => 'n',
            'dddd'  => 'l',
            'ddd'   => 'D',
            'dd'    => 'd',
            'd'     => 'j',
            'ss'    => 's',
            '.s'    => ''
        ),
        '24H' => array(
            'hh' => 'H',
            'h'  => 'G'
        ),
        '12H' => array(
            'hh' => 'h',
            'h'  => 'G'
        )
    );

    /** @var DateTime Standardized base date for the document's date/time values. */
    private static $base_date;

    /** @var bool Is the gmp_gcd method available for usage? Cached value. */
    private static $gmp_gcd_available = false;

    /** @var string Decimal separator character for output of locally formatted values. */
    private $decimal_separator;

    /** @var string Thousands separator character for output of locally formatted values. */
    private $thousand_separator;

    /** @var string Currency character for output of locally formatted values. */
    private $currency_code;

    /** @var SharedStringsConfiguration Configuration of shared strings handling. */
    private $shared_strings_configuration = null;

    /** @var bool Do not format date/time values and return DateTime objects instead. Default false. */
    private $return_date_time_objects;

    /** @var bool Output XLSX-style column names instead of numeric column identifiers. Default false. */
    private $output_column_names;

    /** @var bool Do not consider empty cell values in output. Default false. */
    private $skip_empty_cells;

    /** @var string Full path of the temporary directory that is going to be used to store unzipped files. */
    private $temp_dir;

    /** @var array Temporary files created while reading the document. */
    private $temp_files = array();

    /** @var RelationshipData File paths and -identifiers to all relevant parts of the read XLSX file */
    private $relationship_data;

    /** @var array Data about separate sheets in the file. */
    private $sheets = false;

    /** @var SharedStrings Shared strings handler. */
    private $shared_strings;

    /** @var array Container for cell value style data. */
    private $styles = array();

    /** @var array List of custom formats defined by the current XLSX file; array key = format index */
    private $formats = array();

    /** @var array List of custom formats defined by the user; array key = format index */
    private $customized_formats = array();

    /** @var string Format to use when outputting dates, regardless of originally set formatting.
     *              (Note: Will also be used if the original formatting omits time information, but the data value contains time information.) */
    private $enforced_date_format;

    /** @var string Format to use when outputting time information, regardless of originally set formatting. */
    private $enforced_time_format;

    /** @var string Format to use when outputting datetime values, regardless of originally set formatting. */
    private $enforced_datetime_format;

    /** @var array Cache for already processed format strings. */
    private $parsed_format_cache = array();

    /** @var string Path to the current worksheet XML file. */
    private $worksheet_path = false;

    /** @var OoxmlReader XML reader object for the current worksheet XML file. */
    private $worksheet_reader = false;

    /** @var bool Internal storage for the result of the valid() method related to the Iterator interface. */
    private $valid = false;

    /** @var bool Whether the reader is currently looking at an element within a <row> node. */
    private $row_open = false;

    /** @var int Current row number in the file. */
    private $row_number = 0;

    /** @var bool|array Contents of last read row. */
    private $current_row = false;

    /**
     * @param array $options Reader configuration; Permitted values:
     *      - TempDir (string)
     *          Path to directory to write temporary work files to
     *      - ReturnDateTimeObjects (bool)
     *          If true, date/time data will be returned as PHP DateTime objects.
     *          Otherwise, they will be returned as strings.
     *      - SkipEmptyCells (bool)
     *          If true, row content will not contain empty cells
     *      - SharedStringsConfiguration (SharedStringsConfiguration)
     *          Configuration options to control shared string reading and caching behaviour
     *
     * @throws Exception
     * @throws RuntimeException
     */
    public function __construct(array $options = null)
    {
        if (!isset($options['TempDir'])) {
            $options['TempDir'] = null;
        }
        $this->initTempDir($options['TempDir']);

        if (!empty($options['SharedStringsConfiguration'])) {
            $this->shared_strings_configuration = $options['SharedStringsConfiguration'];
        }
        if (!empty($options['CustomFormats'])) {
            $this->initCustomFormats($options['CustomFormats']);
        }
        if (!empty($options['ForceDateFormat'])) {
            $this->enforced_date_format = $options['ForceDateFormat'];
        }
        if (!empty($options['ForceTimeFormat'])) {
            $this->enforced_time_format = $options['ForceTimeFormat'];
        }
        if (!empty($options['ForceDateTimeFormat'])) {
            $this->enforced_datetime_format = $options['ForceDateTimeFormat'];
        }

        $this->skip_empty_cells = !empty($options['SkipEmptyCells']);
        $this->return_date_time_objects = !empty($options['ReturnDateTimeObjects']);
        $this->output_column_names = !empty($options['OutputColumnNames']);

        $this->initBaseDate();
        $this->initLocale();

        self::$gmp_gcd_available = function_exists('gmp_gcd');
    }

    /**
     * Open the given file and prepare everything for the reading of data.
     *
     * @param   string  $file_path
     *
     * @throws  Exception
     */
    public function open($file_path)
    {
        if (!is_readable($file_path)) {
            throw new RuntimeException('XLSXReader: File not readable (' . $file_path . ')');
        }

        if (!mkdir($this->temp_dir, 0777, true) || !file_exists($this->temp_dir)) {
            throw new RuntimeException(
                'XLSXReader: Could neither create nor confirm existance of temporary directory (' . $this->temp_dir . ')'
            );
        }

        $zip = new ZipArchive;
        $status = $zip->open($file_path);
        if ($status !== true) {
            throw new RuntimeException('XLSXReader: File not readable (' . $file_path . ') (Error ' . $status . ')');
        }

        $this->relationship_data = new RelationshipData($zip);
        $this->initWorkbookData($zip);
        $this->initWorksheets($zip);
        $this->initSharedStrings($zip, $this->shared_strings_configuration);
        $this->initStyles($zip);

        $zip->close();
    }

    /**
     * Free all connected resources.
     */
    public function close()
    {
        if ($this->worksheet_reader && $this->worksheet_reader instanceof OoxmlReader) {
            $this->worksheet_reader->close();
            $this->worksheet_reader = null;
        }

        if ($this->shared_strings && $this->shared_strings instanceof SharedStrings) {
            // Closing the shared string handler will also close all still opened shared string temporary work files.
            $this->shared_strings->close();
            $this->shared_strings = null;
        }

        $this->deleteTempfiles();

        $this->worksheet_path = null;
    }

    /**
     * Set the decimal separator to use for the output of locale-oriented formatted values
     *
     * @param string $new_character
     */
    public function setDecimalSeparator($new_character)
    {
        if (!is_string($new_character)) {
            throw new InvalidArgumentException('Given argument is not a string.');
        }
        $this->decimal_separator = $new_character;
    }

    /**
     * Set the thousands separator to use for the output of locale-oriented formatted values
     *
     * @param string $new_character
     */
    public function setThousandsSeparator($new_character)
    {
        if (!is_string($new_character)) {
            throw new InvalidArgumentException('Given argument is not a string.');
        }
        $this->thousand_separator = $new_character;
    }

    /**
     * Set the currency character code to use for the output of locale-oriented formatted values
     *
     * @param string $new_character
     */
    public function setCurrencyCode($new_character)
    {
        if (!is_string($new_character)) {
            throw new InvalidArgumentException('Given argument is not a string.');
        }
        $this->currency_code = $new_character;
    }

    /**
     * Retrieves an array with information about sheets in the current file
     *
     * @return array List of sheets (key is sheet index, value is of type Worksheet). Sheet's index starts with 0.
     */
    public function getSheets()
    {
        return array_values($this->sheets);
    }

    /**
     * Changes the current sheet in the file to the sheet with the given index.
     *
     * @param   int     $sheet_index
     *
     * @return  bool    True if sheet was successfully changed, false otherwise.
     */
    public function changeSheet($sheet_index)
    {
        $sheets = $this->getSheets(); // Note: Realigns indexes to an auto increment.
        if (!isset($sheets[$sheet_index])) {
            return false;
        }
        /** @var Worksheet $target_sheet */
        $target_sheet = $sheets[$sheet_index];

        // The path to the target worksheet file can be obtained via the relationship id reference.
        $target_relationship_id = $target_sheet->getRelationshipId();
        /** @var RelationshipElement $relationship_worksheet */
        foreach ($this->relationship_data->getWorksheets() as $relationship_worksheet) {
            if ($relationship_worksheet->getId() === $target_relationship_id) {
                $worksheet_path = $relationship_worksheet->getAccessPath();
                break;
            }
        }
        if (!isset($worksheet_path) || !is_readable($worksheet_path)) {
            return false;
        }

        // Initialize the determined target sheet as the new current sheet
        $this->worksheet_path = $worksheet_path;
        $this->rewind();
        return true;
    }

    // !Iterator interface methods

    /**
     * Rewind the Iterator to the first element.
     * Similar to the reset() function for arrays in PHP.
     */
    public function rewind()
    {
        if ($this->worksheet_reader instanceof OoxmlReader) {
            $this->worksheet_reader->close();
        } else {
            $this->worksheet_reader = new OoxmlReader();
            $this->worksheet_reader->setDefaultNamespaceIdentifierElements(OoxmlReader::NS_XLSX_MAIN);
            $this->worksheet_reader->setDefaultNamespaceIdentifierAttributes(OoxmlReader::NS_NONE);
        }

        $this->worksheet_reader->open($this->worksheet_path);

        $this->valid = true;
        $this->row_open = false;
        $this->current_row = false;
        $this->row_number = 0;
    }

    /**
     * Return the current element.
     *
     * @return mixed current element from the collection
     *
     * @throws Exception
     */
    public function current()
    {
        if ($this->row_number === 0 && $this->current_row === false) {
            $this->next();
        }

        return self::adjustRowOutput($this->current_row);
    }

    /**
     * Move forward to next element.
     *
     * @return array|false
     *
     * @throws Exception
     */
    public function next()
    {
        $this->row_number++;

        $this->current_row = array();

        // Walk through the document until the beginning of the first spreadsheet row.
        if (!$this->row_open) {
            while ($this->valid = $this->worksheet_reader->read()) {
                if (!$this->worksheet_reader->matchesElement('row')) {
                    continue;
                }

                $this->row_open = true;

                /* Getting the row spanning area (stored as e.g., 1:12)
                 * so that the last cells will be present, even if empty. */
                $row_spans = $this->worksheet_reader->getAttributeNsId('spans');

                if ($row_spans) {
                    $row_spans = explode(':', $row_spans);
                    $current_row_column_count = $row_spans[1];
                } else {
                    $current_row_column_count = 0;
                }

                // If configured: Return empty strings for empty values
                if ($current_row_column_count > 0 && !$this->skip_empty_cells) {
                    $this->current_row = array_fill(0, $current_row_column_count, '');
                }

                // Do not read further than here if the current 'row' node is not the one to be read
                if ((int) $this->worksheet_reader->getAttributeNsId('r') !== $this->row_number) {
                    return self::adjustRowOutput($this->current_row);
                }
                break;
            }

            // No (further) rows found for reading.
            if (!$this->row_open) {
                return array();
            }
        }

        // Do not read further than here if the current 'row' node is not the one to be read
        if ((int) $this->worksheet_reader->getAttributeNsId('r') !== $this->row_number) {
            $row_spans = $this->worksheet_reader->getAttributeNsId('spans');
            if ($row_spans) {
                $row_spans = explode(':', $row_spans);
                $current_row_column_count = $row_spans[1];
            } else {
                $current_row_column_count = 0;
            }
            if ($current_row_column_count > 0 && !$this->skip_empty_cells) {
                $this->current_row = array_fill(0, $current_row_column_count, '');
            }
            return self::adjustRowOutput($this->current_row);
        }

        // Variables for empty cell handling.
        $max_index = 0;
        $cell_count = 0;
        $last_cell_index = -1;

        // Pre-loop-declarations. Will be filled by one loop iteration, then read in another.
        $style_id = null;
        $cell_index = null;
        $cell_has_shared_string = false;

        while ($this->valid = $this->worksheet_reader->read()) {
            if (!$this->worksheet_reader->matchesNamespace(OoxmlReader::NS_XLSX_MAIN)) {
                continue;
            }
            switch ($this->worksheet_reader->localName) {
                // </row> tag: Finish row reading.
                case 'row':
                    if ($this->worksheet_reader->isClosingTag()) {
                        $this->row_open = false;
                        break 2;
                    }
                    break;

                // <c> tag: Read cell metadata, such as styling of formatting information.
                case 'c':
                    if ($this->worksheet_reader->isClosingTag()) {
                        continue 2;
                    }

                    $cell_count++;

                    // Get the cell index via the "r" attribute and update max_index.
                    $cell_index = $this->worksheet_reader->getAttributeNsId('r');
                    if ($cell_index) {
                        $letter = preg_replace('{[^[:alpha:]]}S', '', $cell_index);
                        $cell_index = self::indexFromColumnLetter($letter);
                    } else {
                        // No "r" attribute available; Just position this cell to the right of the last one.
                        $cell_index = $last_cell_index + 1;
                    }
                    $last_cell_index = $cell_index;
                    if ($cell_index > $max_index) {
                        $max_index = $cell_index;
                    }

                    // Determine cell styling/formatting.
                    $cell_type = $this->worksheet_reader->getAttributeNsId('t');
                    $cell_has_shared_string = $cell_type === 's'; // s = shared string
                    $style_id = (int) $this->worksheet_reader->getAttributeNsId('s');

                    // If configured: Return empty strings for empty values.
                    if (!$this->skip_empty_cells) {
                        $this->current_row[$cell_index] = '';
                    }
                    break;

                // <v> or <is> tag: Read and store cell value according to current styling/formatting.
                case 'v':
                case 'is':
                    if ($this->worksheet_reader->isClosingTag()) {
                        continue 2;
                    }

                    $value = $this->worksheet_reader->readString();

                    if ($cell_has_shared_string) {
                        $value = $this->shared_strings->getSharedString($value);
                    }

                    // Skip empty values when specified as early as possible
                    if ($value === '' && $this->skip_empty_cells) {
                        break;
                    }

                    // Format value if necessary
                    if ($value !== '' && $style_id && isset($this->styles[$style_id])) {
                        $value = $this->formatValue($value, $style_id);
                    } elseif ($value) {
                        $value = $this->generalFormat($value);
                    }

                    $this->current_row[$cell_index] = $value;
                    break;

                default:
                    // nop
                    break;
            }
        }

        /* If configured: Return empty strings for empty values.
         * Only empty cells inbetween and on the left side are added. */
        if (($max_index + 1 > $cell_count) && !$this->skip_empty_cells) {
            $this->current_row += array_fill(0, $max_index + 1, '');
            ksort($this->current_row);
        }

        if (empty($this->current_row) && $this->skip_empty_cells) {
            $this->current_row[] = null;
        }

        return self::adjustRowOutput($this->current_row);
    }

    /**
     * Return the identifying key of the current element.
     *
     * @return mixed either an integer or a string
     */
    public function key()
    {
        return $this->row_number;
    }

    /**
     * Check if there is a current element after calls to rewind() or next().
     * Used to check if we've iterated to the end of the collection.
     *
     * @return boolean FALSE if there's nothing more to iterate over
     */
    public function valid()
    {
        return $this->valid;
    }

    // !Countable interface method

    /**
     * Ostensibly should return the count of the contained items but this just returns the number
     * of rows read so far. It's not really correct but at least coherent.
     */
    public function count()
    {
        return $this->row_number;
    }

    /**
     * Takes the column letter and converts it to a numerical index (0-based)
     *
     * @param   string  $letter Letter(s) to convert
     * @return  mixed   Numeric index (0-based) or boolean false if it cannot be calculated
     */
    public static function indexFromColumnLetter($letter)
    {
        $letter = strtoupper($letter);
        $result = 0;
        for ($i = strlen($letter) - 1, $j = 0; $i >= 0; $i--, $j++) {
            $ord = ord($letter[$i]) - 64;
            if ($ord > 26) {
                // This does not seem to be a letter. Someone must have given us an invalid value.
                return false;
            }
            $result += $ord * (26 ** $j);
        }

        return $result - 1;
    }

    /**
     * Converts the given column index to an XLSX-style [A-Z] column identifier string.
     *
     * @param   int $index
     * @return  string
     */
    public static function columnLetterFromIndex($index)
    {
        $dividend = $index + 1; // Internal counting starts at 0; For easy calculation, it needs to start at 1.
        $output_string = '';
        while ($dividend > 0) {
            $modulo = ($dividend - 1) % 26;
            $output_string = chr($modulo + 65) . $output_string;
            $dividend = floor(($dividend - $modulo) / 26);
        }
        return $output_string;
    }

    /**
     * Helper function for greatest common divisor calculation in case GMP extension is not enabled.
     *
     * @param   int $int_1
     * @param   int $int_2
     * @return  int Greatest common divisor
     */
    private static function GCD($int_1, $int_2)
    {
        $int_1 = (int) abs($int_1);
        $int_2 = (int) abs($int_2);

        if ($int_1 + $int_2 === 0) {
            return 0;
        }

        $divisor = 1;
        while ($int_1 > 0) {
            $divisor = $int_1;
            $int_1 = $int_2 % $int_1;
            $int_2 = $divisor;
        }

        return $divisor;
    }

    /**
     * If configured, replaces numeric column identifiers in output array with XLSX-style [A-Z] column identifiers.
     * If not configured, returns the input array unchanged.
     *
     * @param   array $column_values
     * @return  array
     */
    private function adjustRowOutput($column_values)
    {
        if (!$this->output_column_names) {
            // Column names not desired in output; Nothing to do here.
            return $column_values;
        }

        $column_values_with_keys = array();
        foreach ($column_values as $k => $v) {
            $column_values_with_keys[self::columnLetterFromIndex($k)] = $v;
        }

        return $column_values_with_keys;
    }

    /**
     * Formats the value according to the index.
     *
     * @param   string  $value
     * @param   int     $format_index
     * @return  string
     *
     * @throws  Exception
     */
    private function formatValue($value, $format_index)
    {
        if (!is_numeric($value)) {
            // Only numeric values are formatted.
            return $value;
        }

        if (isset($this->styles[$format_index]) && ($this->styles[$format_index] !== false)) {
            $format_index = $this->styles[$format_index];
        } else {
            // Invalid format_index or the style was explicitly declared as "don't format anything".
            return $value;
        }

        if ($format_index === 0) {
            // Special case for the "General" format
            return $this->generalFormat($value);
        }

        $format = array();
        if (isset($this->parsed_format_cache[$format_index])) {
            $format = $this->parsed_format_cache[$format_index];
        }

        if (!$format) {
            $format = array(
                'Code'      => false,
                'Type'      => false,
                'Scale'     => 1,
                'Thousands' => false,
                'Currency'  => false
            );

            if (array_key_exists($format_index, $this->customized_formats)) {
                $format['Code'] = $this->customized_formats[$format_index];
            } elseif (array_key_exists($format_index, self::BUILTIN_FORMATS)) {
                $format['Code'] = self::BUILTIN_FORMATS[$format_index];
            } elseif (isset($this->formats[$format_index])) {
                $format['Code'] = $this->formats[$format_index];
            }

            // Format code found, now parsing the format
            if ($format['Code']) {
                $sections = explode(';', $format['Code']);
                $format['Code'] = $sections[0];

                switch (count($sections)) {
                    case 2:
                        if ($value < 0) {
                            $format['Code'] = $sections[1];
                        }
                        break;
                    case 3:
                    case 4:
                        if ($value < 0) {
                            $format['Code'] = $sections[1];
                        } elseif ($value === 0) {
                            $format['Code'] = $sections[2];
                        }
                        break;
                    default:
                        // nop
                        break;
                }
            }

            // Stripping colors
            $format['Code'] = trim(preg_replace('{^\[[[:alpha:]]+\]}i', '', $format['Code']));

            // Percentages
            if (substr($format['Code'], -1) === '%') {
                $format['Type'] = 'Percentage';
            } elseif (preg_match('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])*[hmsdy]}i', $format['Code'])) {
                $format['Type'] = 'DateTime';

                $format['Code'] = trim(preg_replace('{^(\[\$[[:alpha:]]*-[0-9A-F]*\])}i', '', $format['Code']));
                $format['Code'] = strtolower($format['Code']);

                $format['Code'] = strtr($format['Code'], self::DATE_REPLACEMENTS['All']);
                if (strpos($format['Code'], 'A') === false) {
                    $format['Code'] = strtr($format['Code'], self::DATE_REPLACEMENTS['24H']);
                } else {
                    $format['Code'] = strtr($format['Code'], self::DATE_REPLACEMENTS['12H']);
                }
            } elseif ($format['Code'] === '[$eUR ]#,##0.00_-') {
                $format['Type'] = 'Euro';
            } else {
                // Removing skipped characters
                $format['Code'] = preg_replace('{_.}', '', $format['Code']);
                // Removing unnecessary escaping
                $format['Code'] = preg_replace("{\\\\}", '', $format['Code']);
                // Removing string quotes
                $format['Code'] = str_replace(array('"', '*'), '', $format['Code']);
                // Removing thousands separator
                if (strpos($format['Code'], '0,0') !== false || strpos($format['Code'], '#,#') !== false) {
                    $format['Thousands'] = true;
                }
                $format['Code'] = str_replace(array('0,0', '#,#'), array('00', '##'), $format['Code']);

                // Scaling (Commas indicate the power)
                $scale = 1;
                $matches = array();
                if (preg_match('{(0|#)(,+)}', $format['Code'], $matches)) {
                    $scale = 1000 ** strlen($matches[2]);
                    // Removing the commas
                    $format['Code'] = preg_replace(array('{0,+}', '{#,+}'), array('0', '#'), $format['Code']);
                }

                $format['Scale'] = $scale;

                if (preg_match('{#?.*\?\/\?}', $format['Code'])) {
                    $format['Type'] = 'Fraction';
                } else {
                    $format['Code'] = str_replace('#', '', $format['Code']);

                    $matches = array();
                    if (preg_match('{(0+)(\.?)(0*)}', preg_replace('{\[[^\]]+\]}', '', $format['Code']), $matches)) {
                        $integer = $matches[1];
                        $decimal_point = $matches[2];
                        $decimals = $matches[3];

                        $format['MinWidth'] = strlen($integer) + strlen($decimal_point) + strlen($decimals);
                        $format['Decimals'] = $decimals;
                        $format['Precision'] = strlen($format['Decimals']);
                        $format['Pattern'] = '%0' . $format['MinWidth'] . '.' . $format['Precision'] . 'f';
                    }
                }

                $matches = array();
                if (preg_match('{\[\$(.*)\]}u', $format['Code'], $matches)) {
                    // Format contains a currency code (Syntax: [$<Currency String>-<language info>])
                    $curr_code = explode('-', $matches[1]);
                    if (isset($curr_code[0])) {
                        $curr_code = $curr_code[0];
                    } else {
                        $curr_code = $this->currency_code;
                    }
                    $format['Currency'] = $curr_code;
                }
                $format['Code'] = trim($format['Code']);
            }
            $this->parsed_format_cache[$format_index] = $format;
        }

        // Applying format to value
        if ($format) {
            if ($format['Code'] === '@') {
                return (string) $value;
            }

            if ($format['Type'] === 'Percentage') {
                // Percentages
                if ($format['Code'] === '0%') {
                    $value = round(100 * $value, 0) . '%';
                } else {
                    $value = sprintf('%.2f%%', round(100 * $value, 2));
                }
            } elseif ($format['Type'] === 'DateTime') {
                // Dates and times
                $days = (int) $value;
                // Correcting for Feb 29, 1900
                if ($days > 60) {
                    $days--;
                }

                // At this point time is a fraction of a day
                $time = ($value - (int) $value);
                $seconds = 0;
                if ($time) {
                    // Here time is converted to seconds
                    // Workaround against precision loss: set low precision will round up milliseconds
                    $seconds = (int) round($time * 86400, 0);
                }

                $original_value = $value;
                $value = clone self::$base_date;
                if ($original_value < 0) {
                    // Negative value, subtract interval
                    $days = abs($days) + 1;
                    $seconds = abs($seconds);
                    $value->sub(new DateInterval('P' . $days . 'D' . ($seconds ? 'T' . $seconds . 'S' : '')));
                } else {
                    // Positive value, add interval
                    $value->add(new DateInterval('P' . $days . 'D' . ($seconds ? 'T' . $seconds . 'S' : '')));
                }

                if (!$this->return_date_time_objects) {
                    // Determine if the format is a date/time/datetime format and apply enforced formatting accordingly
                    $contains_date = preg_match('#[DdFjlmMnoStwWmYyz]#u', $format['Code']);
                    $contains_time = preg_match('#[aABgGhHisuv]#u', $format['Code']);
                    if ($contains_date) {
                        if ($contains_time) {
                            if ($this->enforced_datetime_format) {
                                $value = $value->format($this->enforced_datetime_format);
                            }
                        } else if ($this->enforced_date_format) {
                            $value = $value->format($this->enforced_date_format);
                        }
                    } else if ($this->enforced_time_format) {
                        $value = $value->format($this->enforced_time_format);
                    }

                    if ($value instanceof DateTime) {
                        // No format enforcement for this value type found. Format as declared.
                        $value = $value->format($format['Code']);
                    }
                } // else: A DateTime object is returned
            } elseif ($format['Type'] === 'Euro') {
                $value = 'EUR ' . sprintf('%1.2f', $value);
            } else {
                // Fractional numbers; We get "0.25" and have to turn that into "1/4".
                if ($format['Type'] === 'Fraction' && ($value != (int) $value)) {
                    // Split fraction from integer value (2.25 => 2 and 0.25)
                    $integer = floor(abs($value));
                    $decimal = fmod(abs($value), 1);

                    // Turn fraction into non-decimal value (0.25 => 25)
                    $decimal *= 10 ** (strlen($decimal) - 2);

                    // Obtain biggest divisor for the fraction part (25 => 100 => 25/100)
                    $decimal_divisor = 10 ** strlen($decimal);

                    // Determine greatest common divisor for fraction optimization (so that 25/100 => 1/4)
                    if (self::$gmp_gcd_available) {
                        $gcd = gmp_strval(gmp_gcd($decimal, $decimal_divisor));
                    } else {
                        $gcd = self::GCD($decimal, $decimal_divisor);
                    }

                    // Determine fraction parts (1 and 4 => 1/4)
                    $adj_decimal = $decimal / $gcd;
                    $adj_decimal_divisor = $decimal_divisor / $gcd;

                    if (   strpos($format['Code'], '0') !== false
                        || strpos($format['Code'], '#') !== false
                        || strpos($format['Code'], '? ?') === 0
                    ) {
                        // Extract whole values from fraction (2.25 => "2 1/4")
                        $value = ($value < 0 ? '-' : '') .
                            ($integer ? $integer . ' ' : '') .
                            $adj_decimal . '/' .
                            $adj_decimal_divisor;
                    } else {
                        // Show entire value as fraction (2.25 => "9/4")
                        $adj_decimal += $integer * $adj_decimal_divisor;
                        $value = ($value < 0 ? '-' : '') .
                            $adj_decimal . '/' .
                            $adj_decimal_divisor;
                    }
                } else {
                    // Scaling
                    $value /= $format['Scale'];
                    if (!empty($format['MinWidth']) && $format['Decimals']) {
                        if ($format['Thousands']) {
                            $value = number_format($value, $format['Precision'],
                                $this->decimal_separator, $this->thousand_separator);
                        } else {
                            $value = sprintf($format['Pattern'], $value);
                        }
                        $format_code = preg_replace('{\[\$.*\]}', '', $format['Code']);
                        $value = preg_replace('{(0+)(\.?)(0*)}', $value, $format_code);
                    }
                }

                // Currency/Accounting
                if ($format['Currency']) {
                    $value = preg_replace('{\[\$.*\]}u', $format['Currency'], $value);
                }
            }
        }

        return $value;
    }

    /**
     * Attempts to approximate Excel's "general" format.
     *
     * @param   mixed   $value
     * @return  mixed
     */
    private function generalFormat($value)
    {
        if (is_numeric($value)) {
            $value = (float) $value;
        }

        return $value;
    }

    /**
     * Check and set TempDir to use for file operations.
     * A new folder will be created within the given directory, which will contain all work files,
     * and which will be cleaned up after the process have finished.
     *
     * @param string|null $base_temp_dir
     */
    private function initTempDir($base_temp_dir) {
        if ($base_temp_dir === null) {
            $base_temp_dir = sys_get_temp_dir();
        }
        if (!is_writable($base_temp_dir)) {
            throw new RuntimeException('XLSXReader: Provided temporary directory (' . $base_temp_dir . ') is not writable');
        }
        $base_temp_dir = rtrim($base_temp_dir, DIRECTORY_SEPARATOR);
        /** @noinspection NonSecureUniqidUsageInspection */
        $this->temp_dir = $base_temp_dir . DIRECTORY_SEPARATOR . uniqid() . DIRECTORY_SEPARATOR;
    }

    /**
     * Set base date for calculation of retrieved date/time data.
     *
     * @throws Exception
     */
    private function initBaseDate() {
        self::$base_date = new DateTime();
        self::$base_date->setTimezone(new DateTimeZone('UTC'));
        self::$base_date->setDate(1900, 1, 0);
        self::$base_date->setTime(0, 0, 0);
    }

    /**
     * Pre-fill locale related data using current system locale.
     */
    private function initLocale() {
        $locale = localeconv();
        $this->decimal_separator = $locale['decimal_point'];
        $this->thousand_separator = $locale['thousands_sep'];
        $this->currency_code = $locale['int_curr_symbol'];
    }

    /**
     * Read general workbook information from the given zip into memory.
     *
     * @param   ZipArchive  $zip
     *
     * @throws  Exception
     */
    private function initWorkbookData(ZipArchive $zip)
    {
        $workbook = $this->relationship_data->getWorkbook();
        if (!$workbook) {
            throw new Exception('workbook data not found in XLSX file');
        }
        $workbook_xml = $zip->getFromName($workbook->getOriginalPath());

        $this->sheets = array();
        $workbook_reader = new OoxmlReader();
        $workbook_reader->setDefaultNamespaceIdentifierElements(OoxmlReader::NS_XLSX_MAIN);
        $workbook_reader->setDefaultNamespaceIdentifierAttributes(OoxmlReader::NS_NONE);
        $workbook_reader->xml($workbook_xml);
        while ($workbook_reader->read()) {
            if ($workbook_reader->matchesElement('sheet')) {
                // <sheet/> - Read in data about this worksheet.
                $sheet_name = (string) $workbook_reader->getAttributeNsId('name');
                $rel_id = (string) $workbook_reader->getAttributeNsId('id', OoxmlReader::NS_RELATIONSHIPS_DOCUMENTLEVEL);

                $new_sheet = new Worksheet();
                $new_sheet->setName($sheet_name);
                $new_sheet->setRelationshipId($rel_id);

                $sheet_index = str_ireplace('rId', '', $rel_id);
                $this->sheets[$sheet_index] = $new_sheet;
            } elseif ($workbook_reader->matchesElement('sheets') && $workbook_reader->isClosingTag()) {
                // </sheets> - Indicates that the rest of the document is of no further importance to us. Abort.
                break;
            }
        }
        $workbook_reader->close();
    }

    /**
     * Extract worksheet files to temp directory and set the first worksheet as active.
     *
     * @param ZipArchive $zip
     */
    private function initWorksheets(ZipArchive $zip)
    {
        // Sheet order determining value: relative sheet positioning within the document (rId)
        ksort($this->sheets);

        // Extract worksheets to temporary work directory
        foreach ($this->relationship_data->getWorksheets() as $worksheet) {
            /** @var RelationshipElement $worksheet */
            $worksheet_path_zip = $worksheet->getOriginalPath();
            $worksheet_path_conv = str_replace(RelationshipData::ZIP_DIR_SEP, DIRECTORY_SEPARATOR, $worksheet_path_zip);
            $worksheet_path_unzipped = $this->temp_dir . $worksheet_path_conv;
            if (!$zip->extractTo($this->temp_dir, $worksheet_path_zip)) {
                $message = 'XLSXReader: Could not extract file [' . $worksheet_path_zip . '] to directory [' . $this->temp_dir . '].';
                $this->reportZipExtractionFailure($zip, $message);
            }
            $worksheet->setAccessPath($worksheet_path_unzipped);
            $this->temp_files[] = $worksheet_path_unzipped;
        }

        // Set first sheet as current sheet
        if (!$this->changeSheet(0)) {
            throw new RuntimeException('XLSXReader: Sheet cannot be changed.');
        }
    }

    /**
     * Read shared strings data from the given zip into memory as configured via the given configuration object
     * and potentially create temporary work files for easy retrieval of shared string data.
     *
     * @param ZipArchive                 $zip
     * @param SharedStringsConfiguration $shared_strings_configuration Optional, default null
     */
    private function initSharedStrings(
        ZipArchive $zip,
        SharedStringsConfiguration $shared_strings_configuration = null
    ) {
        $shared_strings = $this->relationship_data->getSharedStrings();
        if (count($shared_strings) > 0) {
            /* Currently, documents with multiple shared strings files are not supported.
            *  Only the first shared string file will be used. */
            /** @var RelationshipElement $first_shared_string_element */
            $first_shared_string_element = $shared_strings[0];

            // Determine target directory and path for the extracted file
            $inzip_path = $first_shared_string_element->getOriginalPath();
            $inzip_path_for_outzip = str_replace(RelationshipData::ZIP_DIR_SEP, DIRECTORY_SEPARATOR, $inzip_path);
            $dir_of_extracted_file = $this->temp_dir . dirname($inzip_path_for_outzip) . DIRECTORY_SEPARATOR;
            $filename_of_extracted_file = basename($inzip_path_for_outzip);
            $path_to_extracted_file = $dir_of_extracted_file . $filename_of_extracted_file;

            // Extract file and note it in relevant variables
            if (!$zip->extractTo($this->temp_dir, $inzip_path)) {
                $message = 'XLSXReader: Could not extract file [' . $inzip_path . '] to directory [' . $this->temp_dir . '].';
                $this->reportZipExtractionFailure($zip, $message);
            }

            $first_shared_string_element->setAccessPath($path_to_extracted_file);
            $this->temp_files[] = $path_to_extracted_file;

            // Initialize SharedStrings
            $this->shared_strings = new SharedStrings(
                $dir_of_extracted_file,
                $filename_of_extracted_file,
                $shared_strings_configuration
            );

            // Extend temp_files with files created by SharedStrings
            $this->temp_files = array_merge($this->temp_files, $this->shared_strings->getTempFiles());
        }
    }

    /**
     * Reads and prepares information on styles declared by the document for later usage.
     *
     * @param ZipArchive $zip
     */
    private function initStyles(ZipArchive $zip)
    {
        $styles = $this->relationship_data->getStyles();
        if (count($styles) > 0) {
            /* Currently, documents with multiple styles files are not supported.
            *  Only the first styles file will be used. */
            /** @var RelationshipElement $first_styles_element */
            $first_styles_element = $styles[0];

            $styles_xml = $zip->getFromName($first_styles_element->getOriginalPath());

            // Read cell style definitions and store them in internal variables
            $styles_reader = new OoxmlReader();
            $styles_reader->setDefaultNamespaceIdentifierElements(OoxmlReader::NS_XLSX_MAIN);
            $styles_reader->setDefaultNamespaceIdentifierAttributes(OoxmlReader::NS_NONE);
            $styles_reader->xml($styles_xml);
            $current_scope_is_cell_xfs = false;
            $current_scope_is_num_fmts = false;
            $switchList = array(
                'numFmts' => array('numFmts'),
                'numFmt'  => array('numFmt'),
                'cellXfs' => array('cellXfs'),
                'xf'      => array('xf')
            );
            while ($styles_reader->read()) {
                switch ($styles_reader->matchesOneOfList($switchList)) {
                    // <numFmts><numFmt/></numFmts> - check for number format definitions
                    case 'numFmts':
                        $current_scope_is_num_fmts = !$styles_reader->isClosingTag();
                        break;
                    case 'numFmt':
                        if (!$current_scope_is_num_fmts || $styles_reader->isClosingTag()) {
                            break;
                        }
                        $format_code = (string) $styles_reader->getAttributeNsId('formatCode');
                        $num_fmt_id = (int) $styles_reader->getAttributeNsId('numFmtId');
                        $this->formats[$num_fmt_id] = $format_code;
                        break;

                    // <cellXfs><xf/></cellXfs> - check for format usages
                    case 'cellXfs':
                        $current_scope_is_cell_xfs = !$styles_reader->isClosingTag();
                        break;
                    case 'xf':
                        if (!$current_scope_is_cell_xfs || $styles_reader->isClosingTag()) {
                            break;
                        }

                        // Determine if number formatting is set for this cell.
                        $num_fmt_id = null;
                        if ($styles_reader->getAttributeNsId('numFmtId')) {
                            $applyNumberFormat = $styles_reader->getAttributeNsId('applyNumberFormat');
                            if ($applyNumberFormat === null || $applyNumberFormat === '1' || $applyNumberFormat === 'true') {
                                /* Number formatting is enabled either implicitly ('applyNumberFormat' not given)
                                 * or explicitly ('applyNumberFormat' is a true value). */
                                $num_fmt_id = (int) $styles_reader->getAttributeNsId('numFmtId');
                            }
                        }

                        // Determine and store correct formatting style.
                        if ($num_fmt_id !== null) {
                            // Number formatting has been enabled for this format.
                            // If format ID >= 164, it is a custom format and should be read from styleSheet\numFmts
                            $this->styles[] = $num_fmt_id;
                        } else if ($styles_reader->getAttributeNsId('quotePrefix')) {
                            // "quotePrefix" automatically preceeds the cell content with a ' symbol. This enforces a text format.
                            $this->styles[] = false; // false = "Do not format anything".
                        } else {
                            $this->styles[] = 0; // 0 = "General" format
                        }
                        break;
                }
            }
            $styles_reader->close();
        }
    }

    /**
     * @param   array $custom_formats
     * @return  void
     */
    private function initCustomFormats(array $custom_formats)
    {
        foreach ($custom_formats as $format_index => $format) {
            if (array_key_exists($format_index, self::BUILTIN_FORMATS) !== false) {
                $this->customized_formats[$format_index] = $format;
            }
        }
    }

    /**
     * Delete all registered temporary work files and -directories.
     */
    private function deleteTempfiles() {
        foreach ($this->temp_files as $temp_file) {
            @unlink($temp_file);
        }

        // Better safe than sorry - shouldn't try deleting '.' or '/', or '..'.
        if (strlen($this->temp_dir) > 2) {
            @rmdir($this->temp_dir . 'xl' . DIRECTORY_SEPARATOR . 'worksheets');
            @rmdir($this->temp_dir . 'xl');
            @rmdir($this->temp_dir);
        }
    }

    /**
     * Gather data on zip extractTo() fault and throw an appropriate Exception.
     *
     * @param   ZipArchive  $zip
     * @param   string      $message    Optional error message to prefix the error details with.
     *
     * @throws  RuntimeException
     */
    private function reportZipExtractionFailure($zip, $message = '')
    {
        $status_code = $zip->status;
        $status_message = $zip->getStatusString();
        if ($status_code || $status_message) {
            $message .= ' Status from ZipArchive:';
            if ($status_code) {
                $message .= ' Code [' . $status_code . '];';
            }
            if ($status_message) {
                $message .= ' Message [' . $status_message . '];';
            }
        }
        throw new RuntimeException($message);
    }
}
?>
