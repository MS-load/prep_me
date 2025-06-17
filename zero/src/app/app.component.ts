import { Component } from '@angular/core';
import { WishItem } from 'src/shared/models/wishItem';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})

type listFilters =  'all' | 'comp' | 'not_comp'
export class AppComponent {
  title = 'zero to hero';

  items: WishItem[] = [
    new WishItem('Build one ionic app', true),
    new WishItem('Learn Angular'),
    new WishItem('Something to check off'),
  ];

  listFilters: listFilters= 'all';

  newWishText = '';

  visibleItems: WishItem[] = [];

  completedItems: number =
    this.items.filter(({ isComplete }) => isComplete)?.length ?? 0;

  addItem() {
    this.items.push(new WishItem(this.newWishText));
    this.newWishText = '';
  }

  filterChanged(value: listFilters) {
    if (value === 'all') {
      return (this.visibleItems = this.items);
    }
    
    if(value === 'comp'){
      return this.visibleItems = this.items.filter((itm)=> itm.isComplete)
    }

    return this.visibleItems = this.items.filter((itm)=> !itm.isComplete)
  }

  toggleItem(item: WishItem) {
    item.isComplete = !item.isComplete;
    console.log(item);
  }
}
