import { Injectable, signal, computed } from '@angular/core';

export type NotificationType = 'success' | 'error' | 'warning' | 'info';

export interface Notification {
  id: number;
  type: NotificationType;
  message: string;
  duration?: number;
}

/**
 * Serwis powiadomień — centralne zarządzanie komunikatami dla użytkownika
 */
@Injectable({
  providedIn: 'root'
})
export class NotificationService {
  private _notifications = signal<Notification[]>([]);
  private _nextId = 0;

  readonly notifications = this._notifications.asReadonly();
  readonly hasNotifications = computed(() => this._notifications().length > 0);

  /**
   * Wyświetla powiadomienie
   */
  show(type: NotificationType, message: string, duration = 5000): void {
    const notification: Notification = {
      id: ++this._nextId,
      type,
      message,
      duration
    };

    this._notifications.update(list => [...list, notification]);

    if (duration > 0) {
      setTimeout(() => this.dismiss(notification.id), duration);
    }
  }

  success(message: string): void {
    this.show('success', message);
  }

  error(message: string): void {
    this.show('error', message, 8000);
  }

  warning(message: string): void {
    this.show('warning', message);
  }

  info(message: string): void {
    this.show('info', message);
  }

  /**
   * Usuwa powiadomienie
   */
  dismiss(id: number): void {
    this._notifications.update(list => list.filter(n => n.id !== id));
  }

  /**
   * Usuwa wszystkie powiadomienia
   */
  clear(): void {
    this._notifications.set([]);
  }
}
