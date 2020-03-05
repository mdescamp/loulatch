<?php

declare(strict_types=1);

namespace Loulatch\DocxConverterBundle\Controller;

use AppBundle\Entity\Document;
use AppBundle\Form\DocumentType;
use AppBundle\Services\PDFHelper;
use Doctrine\ORM\EntityManagerInterface;
use Doctrine\ORM\OptimisticLockException;
use Doctrine\ORM\ORMException;
use Doctrine\ORM\TransactionRequiredException;
use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\Form\FormInterface;
use Symfony\Component\HttpFoundation\BinaryFileResponse;
use Symfony\Component\HttpFoundation\File\Exception\FileException;
use Symfony\Component\HttpFoundation\File\UploadedFile;
use Symfony\Component\HttpFoundation\JsonResponse;
use Symfony\Component\HttpFoundation\RedirectResponse;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;

class DocumentController extends Controller
{
    /**
     * @param Request $request
     * @return JsonResponse
     */
    public function listDocModelAction(Request $request): JsonResponse
    {
        $category = $request->get('category');
        $documents = $this->em->getRepository(Document::class)
            ->findByAgenceCategory($this->agenceId, $category);

        $options = $this->renderView('_model-doc-select.html.twig', ['documents' => $documents]);

        return $this->json($options);
    }

    private $em;
    private $pdfHelper;

    public function __construct(EntityManagerInterface $em, PDFHelper $pdfHelper)
    {
        $this->em = $em;
        $this->pdfHelper = $pdfHelper;
    }

    /**
     * @param Request  $request
     * @param Document $document
     * @return RedirectResponse
     */
    public function deleteAction(Request $request, Document $document): RedirectResponse
    {
        $form = $this->createDeleteForm($document);
        $form->handleRequest($request);

        if ($form->isSubmitted() && $form->isValid()) {
            $em = $this->getDoctrine()->getManager();
            if (false === unlink($this->getParameter('document_directory') . $document->getFilePath())) {
                throw new RuntimeException(sprintf('Le fichier  \'%s\' n\'a pas pu être supprimé', $document->getFilePath()));
            }
            $em->remove($document);
            $em->flush();
        }

        return $this->redirectToRoute('document_index');
    }

    /**
     * @param Request  $request
     * @param Document $document
     * @return Response
     */
    public function editAction(Request $request, Document $document): Response
    {
        $deleteForm = $this->createDeleteForm($document);
        $form = $this->createForm(DocumentType::class, $document, [
            'mode' => 'edit',
        ]);

        $form->handleRequest($request);
        if ($form->isSubmitted() && $form->isValid()) {
            $this->saveFilePath($form, $document);
            $this->em->persist($document);
            $this->em->flush();
            $this->addFlash('success', 'Document mis à jour.');
            return $this->redirectToRoute('document_index');
        }

        return $this->render('AppBundle:Document:edit.html.twig', [
            'form' => $form->createView(),
            'deleteForm' => $deleteForm->createView(),
            'document' => $document,
        ]);
    }

    /**
     * @return Response
     */
    public function indexAction(): Response
    {
        /** @var Document[] $documents */
        $documents = $this->em->getRepository(Document::class)->findAll();
        return $this->render('AppBundle:Document:index.html.twig', [
            'documents' => $documents,
        ]);
    }

    /**
     * @param Request $request
     * @return Response
     */
    public function newAction(Request $request): Response
    {
        $document = new Document();
        $form = $this->createForm(DocumentType::class, $document, [
            'mode' => 'new',
        ]);

        $form->handleRequest($request);
        if ($form->isSubmitted() && $form->isValid()) {
            $this->saveFilePath($form, $document);
            $this->em->persist($document);
            $this->em->flush();
            return $this->redirectToRoute('document_index');
        }

        return $this->render('AppBundle:Document:new.html.twig', [
            'form' => $form->createView()
        ]);
    }

    /**
     * @param Document $document
     * @return FormInterface
     */
    private function createDeleteForm(Document $document): FormInterface
    {
        return $this->createFormBuilder()
            ->setAction($this->generateUrl('document_delete', ['id' => $document->getId()]))
            ->setMethod('DELETE')
            ->getForm();
    }

    /**
     * @param FormInterface $form
     * @param Document      $document
     */
    private function saveFilePath(FormInterface $form, Document $document): void
    {
        /** @var UploadedFile $filePath */
        $filePath = $form->get('filePath')->getData();
        if ($filePath instanceof UploadedFile) {
            $originalFilename = pathinfo($filePath->getClientOriginalName(), PATHINFO_FILENAME);
            $safeFilename = transliterator_transliterate('Any-Latin; Latin-ASCII; [^A-Za-z0-9_] remove; Lower()', $originalFilename);
            $newFilename = $safeFilename . '-' . uniqid('', true) . '.' . $filePath->guessExtension();
            try {
                $filePath->move($this->getParameter('document_directory'), $newFilename);
            } catch (FileException $e) {
                echo $e->getMessage();
            }
            $document->setFilePath($newFilename);
        } else {
            $document->setFilePath($document->getFilePath());
        }
    }

    public function printAction(Request $request, int $id): BinaryFileResponse
    {
        $document = $this->em->find(Document::class, $request->get('document'));
        $options = !empty($request->get('options')) ? json_decode($request->get('options'), true) : [];

        try {
            $file = $this->pdfHelper->getPdfDocument($document, $id, $request->get('name'), $options);
        } catch (OptimisticLockException $e) {
            echo $e->getMessage();
        } catch (TransactionRequiredException $e) {
            echo $e->getMessage();
        } catch (ORMException $e) {
            echo $e->getMessage();
        } catch (CopyFileException $e) {
            echo $e->getMessage();
        } catch (CreateTemporaryFileException $e) {
            echo $e->getMessage();
        }
        return $this->file($file ?? null);
    }
}
