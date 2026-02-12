import { ComponentFixture, TestBed } from '@angular/core/testing';
import { of } from 'rxjs';
import { DocxEditorComponent } from './docx-editor';
import { FileUpload as FileUploadService } from '../../services/file-upload';

describe('DocxEditorComponent', () => {
  let component: DocxEditorComponent;
  let fixture: ComponentFixture<DocxEditorComponent>;
  let openDocxCalledWith: File | null = null;
  const fileUploadServiceMock = {
    openDocx: (file: File) => {
      openDocxCalledWith = file;
      return of({
        success: true,
        message: 'ok',
        fileName: 'test.docx',
        html: '<p>Test</p>'
      });
    },
    saveDocx: (_request: { fileName?: string; html: string }) => of(new Blob())
  };

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [DocxEditorComponent],
      providers: [{ provide: FileUploadService, useValue: fileUploadServiceMock }]
    }).compileComponents();

    fixture = TestBed.createComponent(DocxEditorComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });

  it('should open docx and set editor content', () => {
    const file = new File(['docx'], 'test.docx');
    component.selectedDocx = file;

    component.openDocx();

    expect(openDocxCalledWith).toBe(file);
    expect(component.editorContent).toBe('<p>Test</p>');
  });
});
