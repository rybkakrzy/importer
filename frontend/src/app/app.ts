import { Component, signal } from '@angular/core';
import { FileUploadComponent } from './components/file-upload/file-upload';

@Component({
  selector: 'app-root',
  imports: [FileUploadComponent],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class App {
  protected readonly title = signal('Importer Parametryzacji');
}
