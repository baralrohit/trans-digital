import { Component } from '@angular/core';
// import { RouterOutlet } from '@angular/router';
import { NavbarComponent } from './navbar/navbar.component';
import { GtfsBuilderComponent } from './gtfs-builder/gtfs-builder.component';

@Component({
  selector: 'app-root',
  imports: [NavbarComponent, GtfsBuilderComponent],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss',
})
export class AppComponent {
  title = 'trans-digital';
}
