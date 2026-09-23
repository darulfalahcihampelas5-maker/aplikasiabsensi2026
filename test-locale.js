import { id } from 'date-fns/locale';
import { format } from 'date-fns';
console.log(format(new Date(), 'EEEE, dd MMMM yyyy', { locale: id }));
