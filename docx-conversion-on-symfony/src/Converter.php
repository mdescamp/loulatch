<?php


use AppBundle\Entity\Document;
use Doctrine\ORM\EntityManager;
use Doctrine\ORM\OptimisticLockException;
use Doctrine\ORM\ORMException;
use Doctrine\ORM\TransactionRequiredException;
use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\TemplateProcessor;
use Symfony\Component\HttpFoundation\Session\SessionInterface;
use Symfony\Component\HttpKernel\KernelInterface;
use Symfony\Component\Process\Process;
use Symfony\Component\Yaml\Exception\ParseException;
use Symfony\Component\Yaml\Yaml;

class Converter
{
//    private $agenceId;
//    public $configPath;
//    private $em;
//    public $modelPath;
//    public $ouputPath;
//    private $session;
//
//    public function __construct()
//    {
//        $this->em = $em;
//        $this->session = $session;
//        $this->agenceId = $session->get('agenceParDefaut');
//        $this->ouputPath = '/tmp/';
//        $this->modelPath = $kernel->getContainer()->getParameter('document_directory');
//        $this->configPath = __DIR__ . '/../Resources/DocxVariables/';
//    }

    /**
     * @param $cmd
     */
    public function checkBinaryExist($cmd): void
    {
        $return = shell_exec(sprintf('which %s', escapeshellarg($cmd)));
        if (empty($return)) {
            throw new RuntimeException(sprintf('Binary \'%s\' not found.', $cmd));
        }
    }

    /**
     * @param array $customVariables
     * @param array $options
     */
    private function checkCustomVariables(array $customVariables, array $options): void
    {
        foreach ($customVariables as $entityName => $configGetters) {
            if (!class_exists('AppBundle\Entity\\' . $entityName)) {
                throw new RuntimeException(sprintf('Class AppBundle\Entity\%s does not exist', $entityName));
            }

            if (!isset($configGetters['config']) || count($configGetters['config']) === 0) {
                throw new RuntimeException(sprintf('No config found for %s entity', $entityName));
            }

            if (!isset($configGetters['config']['find']) && !isset($configGetters['array'])) {
                throw new RuntimeException(sprintf('Key \'find\' in config for %s entity not found', $entityName));
            }

            if (!isset($configGetters['config']['var'])) {
                throw new RuntimeException(sprintf('Key \'var\' in config for %s entity not found', $entityName));
            }

            if ((!isset($configGetters['getters']) || count($configGetters['getters']) === 0) && !isset($configGetters['array'])) {
                throw new RuntimeException(sprintf('No getters find for %s entity', $entityName));
            }

            if (isset($configGetters['getters'])) {
                foreach ($configGetters['getters'] as $getter) {
                    if (!method_exists('AppBundle\Entity\\' . $entityName, 'get' . $getter)) {
                        throw new RuntimeException(
                            sprintf('Method %s for entity AppBundle\Entity\%s does not exist', $getter, $entityName)
                        );
                    }
                }
            }
            if (isset($configGetters['array'])) {
                foreach ($configGetters['array'] as $configGetter) {
                    if (is_array($configGetter) && !isset($options[reset($configGetter)])) {
                        throw new RuntimeException(sprintf('The Entity \'%s\' need \'%s\' option', $entityName, reset($configGetter)));
                    }
                }
            }
        }
    }

    /**
     * @param Document $document
     */
    private function checkModel(Document $document): void
    {
        $this->modelPath .= '/' . $document->getFilePath();
        if (!file_exists($this->modelPath)) {
            throw new RuntimeException(sprintf('File %s does not exist', $this->modelPath));
        }
        if (pathinfo($this->modelPath, PATHINFO_EXTENSION) !== 'docx') {
            throw new RuntimeException(sprintf('File %s is not a docx type', $this->modelPath));
        }
    }

    /**
     * @param array $configs
     * @param int   $id
     * @return array
     * @throws ORMException
     * @throws OptimisticLockException
     * @throws TransactionRequiredException
     */
    private function getAllEntities(array $configs, int $id): array
    {
        $agenceId = $this->agenceId;
        foreach ($configs as $entityName => $entity) {
            if (isset($entity['config']['parent'])) {
                if (($tmp = $this->getParent($entity['config'], $configs, $id)) !== null) {
                    $entities[$entityName] = $tmp;
                }
            } elseif (isset($entity['config']['find']) && is_array($entity['config']['find'])) {
                $entities[$entityName] = $this->em->getRepository('AppBundle:' . $entityName)->findBy([$entity['config']['find'][0] => $id]);
            } elseif (isset($entity['config']['find'])) {
                $entities[$entityName] = $this->em->find('AppBundle:' . $entityName, ${$entity['config']['find']});
            }
        }
        return $entities ?? [];
    }

    /**
     * @param string $doctype
     * @return array
     */
    private function getFileConfigs(string $doctype): array
    {
        try {
            // global is Agence
            $global = Yaml::parseFile($this->configPath . 'global.yml');
            $specifique = Yaml::parseFile($this->configPath . $doctype . '.yml');
        } catch (ParseException $e) {
            throw new RuntimeException('Config file is not a valid yaml');
        }

        if ($global === null || $specifique === null) {
            throw new RuntimeException('Config file is empty');
        }

        return array_merge($global, $specifique);
    }

    /**
     * @param $config
     * @param $customVariables
     * @param $id
     * @return mixed
     */
    private function getParent(array $config, array $customVariables, int $id)
    {
        if (!isset(${$customVariables[$config['parent']]['config']['find']})) {
            // Parent is a child too
            $parentEntity = $this->getParent(
                $customVariables[$config['parent']]['config'],
                $customVariables,
                $id
            );
        } else {
            // Is the last Parent of the all generation
            $parentEntity = $this->em
                ->getRepository('AppBundle:' . $config['parent'])
                ->find(${$customVariables[$config['parent']]['config']['find']});
        }
        if ($parentEntity !== null) {
            return $parentEntity->{'get' . ucfirst($config['find'])}();
        }
        return null;
    }

    /**
     * @param Document $document
     * @param int      $id
     * @param string   $outputName
     * @param array    $options
     * @return string|null
     * @throws CopyFileException
     * @throws CreateTemporaryFileException
     * @throws ORMException
     * @throws OptimisticLockException
     * @throws TransactionRequiredException
     */
    public function getPdfDocument(Document $document, int $id, string $outputName, array $options = []): ?string
    {
        $this->checkModel($document);
        $configs = $this->getFileConfigs($document->getCategory());
        $this->checkCustomVariables($configs, $options);
        $entities = $this->getAllEntities($configs, $id);
        [$simpleVariables, $dateVariables, $arrayVariables] = $this->useAllGetters($entities, $configs, $id);
        $template = $this->remplaceAllVariables($simpleVariables, $dateVariables, $arrayVariables, $id, $options);
        $outputName = $this->transformOutputName($outputName, $simpleVariables, $document);
        $this->saveTmpDocx($template, $outputName);
        return $this->transformDocxToPdf($outputName);
    }

    /**
     * @param array $entities
     * @param       $getter
     * @param       $entityName
     * @return string|string[]
     */
    private function getStrReplaceForXml(array $entities, $getter, $entityName)
    {
        $value = $entities[$entityName]->{'get' . ucfirst($getter)}();
        if (is_float($value)) {
            $value = round($value, 2);
        }
        return str_replace(
            "\n", '</w:t><w:br/><w:t>',
            htmlspecialchars((string)$value)
        );
    }

    /**
     * @param array $result
     * @return array
     */
    private function htmlEncode(array $result): array
    {
        if (!empty($result)) {
            foreach ($result as $key => $value) {
                if (is_array($value)) {
                    foreach ($value as $key2 => $value2) {
                        $result[$key][$key2] = htmlspecialchars($value2);
                    }
                }
            }
        }
        return $result;
    }

    /**
     * @param TemplateProcessor $template
     */
    private function removeVariablesNotFound(TemplateProcessor $template): void
    {
        $variables = $template->getVariables();
        foreach ($variables as $variable) {
            $template->setValue($variable, '');
        }
    }

    /**
     * @param array $allVariables
     * @param array $dateVariables
     * @param array $arrayVariables
     * @param int   $id
     * @param array $options
     * @return TemplateProcessor
     * @throws CopyFileException
     * @throws CreateTemporaryFileException
     */
    public function remplaceAllVariables(array $allVariables, array $dateVariables, array $arrayVariables, int $id, array $options): TemplateProcessor
    {
        $template = new TemplateProcessor($this->modelPath);
        if (!empty($dateVariables)) {
            $this->remplaceDateVariables($dateVariables, $template);
        }
        if (!empty($arrayVariables)) {
            $custom = $this->remplaceArrayVariables($arrayVariables, $id, $template, $options);
            if ($custom !== null) {
                $allVariables = array_merge($allVariables, $custom);
            }
        }
        $template->setValues($allVariables);
        $this->removeVariablesNotFound($template);
        return $template;
    }

    /**
     * @param array             $arrayVariables
     * @param int               $id
     * @param TemplateProcessor $template
     * @param array             $options
     * @return mixed|null
     */
    private function remplaceArrayVariables(array $arrayVariables, int $id, TemplateProcessor $template, array $options)
    {
        $custom = [];
        foreach ($arrayVariables as $entityName => $array) {
            foreach ($array as $arrayName) {
                $entity = 'AppBundle\Models\\' . $entityName . 'DocArray';
                if (is_array($arrayName)) {
                    if (!isset($options[$arrayName[array_key_first($arrayName)]])) {
                        throw new RuntimeException(sprintf('Option %s not found', $arrayName[1]));
                    }
                    $array = (new $entity($this->em, $this->session, $id))->{'get' . ucfirst(array_key_first($arrayName))}($options[$arrayName[array_key_first($arrayName)]]);
                } else {
                    $array = (new $entity($this->em, $this->session, $id))->{'get' . ucfirst($arrayName)}();
                }
                if (isset($array['custom'])) {
                    $custom = $array['custom'];
                    unset($array['custom']);
                }
                if (!empty($array) && isset($template->getVariableCount()[array_key_first($array[0])])) {
                    $template->cloneRowAndSetValues(array_key_first($array[0]), $this->htmlEncode($array));
                }
            }
        }
        return $custom ?? null;
    }

    /**
     * @param array             $dateVariables
     * @param TemplateProcessor $template
     */
    private function remplaceDateVariables(array $dateVariables, TemplateProcessor $template): void
    {
        $variables = $template->getVariables();
        foreach ($variables as $variable) {
            foreach ($dateVariables as $dateVariable => $dateTime) {
                if (strpos($variable, $dateVariable) !== false) {
                    if (preg_match('/\(.*\)/', $variable, $match) === 1) {
                        $template->setValue($variable, $dateTime->format(
                            str_replace(['(', ')'], '', $match[0])
                        ));
                    } else {
                        $template->setValue($variable, $dateTime->format('Y-m-d H:i:s'));
                    }
                }
            }
        }
    }

    /**
     * @param TemplateProcessor $template
     * @param string            $outputName
     */
    private function saveTmpDocx(TemplateProcessor $template, string $outputName): void
    {
        if (file_exists($this->ouputPath . $outputName . '.docx')) {
            unlink($this->ouputPath . $outputName . '.docx');
        }
        $template->saveAs($this->ouputPath . $outputName . '.docx');
    }

    /**
     * @param string $outputName
     * @return string|null
     */
    public function transformDocxToPdf(string $outputName): ?string
    {
        $this->checkBinaryExist('lowriter');
        $command = 'sudo lowriter --convert-to pdf "' . $this->ouputPath . $outputName . '.docx" --outdir "' . $this->ouputPath . '"';
        $process = new Process($command);
        $process->start();
        while ($process->isRunning()) {
            // wait
        }
        if (file_exists($this->ouputPath . $outputName . '.pdf')) {
            return $this->ouputPath . $outputName . '.pdf';
        }
        return null;
    }

    /**
     * @param array $entities
     * @param array $configs
     * @param int   $id
     * @return array
     */
    public function useAllGetters(array $entities, array $configs, int $id): array
    {
        $simpleVariables = [];
        $dateVariables = [];
        $arrayVariables = [];
        foreach ($configs as $entityName => $config) {
            if (isset($config['config']['self'])) {
                $simpleVariables[$config['config']['var'] . '_' . $config['config']['self']] = $id;
            }
            if (isset($config['getters'])) {
                foreach ($config['getters'] as $var => $getter) {
                    if (isset($entities[$entityName])) {
                        if ($entities[$entityName]->{'get' . ucfirst($getter)}() instanceof DateTime) {
                            $dateVariables[$config['config']['var'] . '_' . $var] = $entities[$entityName]->{'get' . ucfirst($getter)}();
                        } else {
                            $simpleVariables[$config['config']['var'] . '_' . $var] = $this->getStrReplaceForXml($entities, $getter, $entityName);
                        }
                    }
                }
            }
            if (isset($config['array'])) {
                $arrayVariables[$entityName] = $config['array'];
            }
        }
        return [$simpleVariables, $dateVariables, $arrayVariables];
    }

    /**
     * @param string   $outputName
     * @param array    $simpleVariables
     * @param Document $document
     * @return string|string[]
     */
    private function transformOutputName(string $outputName, array $simpleVariables, Document $document): string
    {
        if (empty($outputName)) {
            return $document->getName();
        }
        $matches = [];
        preg_match_all('/\$\{(.*?)}/i', $outputName, $matches);
        $matches = array_combine($matches[0], $matches[1]);
        foreach ($matches as $dynamic => $variable) {
            if ($variable === 'DOCUMENT') {
                $outputName = str_replace($dynamic, $document->getName(), $outputName);
            } else {
                $outputName = str_replace($dynamic, $simpleVariables[$variable] ?? '', $outputName);
            }
        }
        return $outputName;
    }
}
