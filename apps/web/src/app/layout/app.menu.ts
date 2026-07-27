import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterModule } from '@angular/router';
import { MenuItem } from 'primeng/api';
import { AppMenuitem } from './app.menuitem';

@Component({
    selector: 'app-menu',
    standalone: true,
    imports: [CommonModule, AppMenuitem, RouterModule],
    template: `<ul class="layout-menu">
        @for (item of model; track item.label) {
            @if (!item.separator) {
                <li app-menuitem [item]="item" [root]="true"></li>
            } @else {
                <li class="menu-separator"></li>
            }
        }
    </ul>`
})
export class AppMenu {
    model: MenuItem[] = [];

    ngOnInit() {
        this.model = [
            {
                label: 'Principal',
                items: [
                    { label: 'Dashboard', icon: 'pi pi-fw pi-home', routerLink: ['/dashboard'] }
                ]
            },
            {
                label: 'Microsoft 365',
                items: [
                    { label: 'Servicios M365', icon: 'pi pi-fw pi-microsoft', routerLink: ['/microsoft'] },
                    { label: 'Calendario', icon: 'pi pi-fw pi-calendar', routerLink: ['/calendar'] },
                    { label: 'Correo', icon: 'pi pi-fw pi-envelope', routerLink: ['/mail'] },
                    { label: 'Documentos', icon: 'pi pi-fw pi-file', routerLink: ['/documents'] },
                    { label: 'Teams', icon: 'pi pi-fw pi-comments', routerLink: ['/teams'] }
                ]
            },
            {
                label: 'Validación',
                items: [
                    { label: 'Test Suite M365', icon: 'pi pi-fw pi-play', routerLink: ['/test-suite'] }
                ]
            }
        ];
    }
}
