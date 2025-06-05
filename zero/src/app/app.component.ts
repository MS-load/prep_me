import { Component } from '@angular/core';
import { WishItem } from 'src/shared/models/wishItem';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'zero to hero';
  items:  WishItem[] =  [
    new WishItem('Build one ionic app', true ),
    new WishItem('Learn Angular'),
    new WishItem('Something to check off')
  ]
}
